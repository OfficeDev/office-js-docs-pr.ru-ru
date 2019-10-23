---
title: Office. Context. Mailbox — набор обязательных элементов 1,4
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 46a73e4911d95310efbe0607b6ba0715238cd6cc
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626987"
---
# <a name="mailbox"></a><span data-ttu-id="86b7a-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="86b7a-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="86b7a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="86b7a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="86b7a-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="86b7a-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="86b7a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-105">Requirements</span></span>

|<span data-ttu-id="86b7a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-106">Requirement</span></span>| <span data-ttu-id="86b7a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-109">1.0</span></span>|
|[<span data-ttu-id="86b7a-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="86b7a-111">Restricted</span></span>|
|[<span data-ttu-id="86b7a-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="86b7a-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="86b7a-114">Members and methods</span></span>

| <span data-ttu-id="86b7a-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="86b7a-115">Member</span></span> | <span data-ttu-id="86b7a-116">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="86b7a-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="86b7a-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="86b7a-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="86b7a-118">Member</span></span> |
| [<span data-ttu-id="86b7a-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="86b7a-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="86b7a-120">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-120">Method</span></span> |
| [<span data-ttu-id="86b7a-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="86b7a-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="86b7a-122">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-122">Method</span></span> |
| [<span data-ttu-id="86b7a-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="86b7a-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="86b7a-124">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-124">Method</span></span> |
| [<span data-ttu-id="86b7a-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="86b7a-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="86b7a-126">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-126">Method</span></span> |
| [<span data-ttu-id="86b7a-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="86b7a-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="86b7a-128">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-128">Method</span></span> |
| [<span data-ttu-id="86b7a-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="86b7a-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="86b7a-130">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-130">Method</span></span> |
| [<span data-ttu-id="86b7a-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="86b7a-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="86b7a-132">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-132">Method</span></span> |
| [<span data-ttu-id="86b7a-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="86b7a-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="86b7a-134">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-134">Method</span></span> |
| [<span data-ttu-id="86b7a-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="86b7a-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="86b7a-136">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-136">Method</span></span> |
| [<span data-ttu-id="86b7a-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="86b7a-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="86b7a-138">Метод</span><span class="sxs-lookup"><span data-stu-id="86b7a-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="86b7a-139">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="86b7a-139">Namespaces</span></span>

<span data-ttu-id="86b7a-140">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="86b7a-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="86b7a-141">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="86b7a-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="86b7a-142">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="86b7a-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="86b7a-143">Members</span><span class="sxs-lookup"><span data-stu-id="86b7a-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="86b7a-144">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="86b7a-144">ewsUrl: String</span></span>

<span data-ttu-id="86b7a-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-147">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="86b7a-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="86b7a-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="86b7a-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="86b7a-150">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="86b7a-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="86b7a-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="86b7a-153">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-153">Type</span></span>

*   <span data-ttu-id="86b7a-154">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="86b7a-155">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-155">Requirements</span></span>

|<span data-ttu-id="86b7a-156">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-156">Requirement</span></span>| <span data-ttu-id="86b7a-157">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-159">1.0</span></span>|
|[<span data-ttu-id="86b7a-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-161">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="86b7a-164">Методы</span><span class="sxs-lookup"><span data-stu-id="86b7a-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="86b7a-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="86b7a-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="86b7a-166">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="86b7a-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-167">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="86b7a-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="86b7a-p104">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-170">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-170">Parameters</span></span>

|<span data-ttu-id="86b7a-171">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-171">Name</span></span>| <span data-ttu-id="86b7a-172">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-172">Type</span></span>| <span data-ttu-id="86b7a-173">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="86b7a-174">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-174">String</span></span>|<span data-ttu-id="86b7a-175">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="86b7a-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="86b7a-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="86b7a-177">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="86b7a-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-178">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-178">Requirements</span></span>

|<span data-ttu-id="86b7a-179">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-179">Requirement</span></span>| <span data-ttu-id="86b7a-180">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="86b7a-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-182">1.3</span><span class="sxs-lookup"><span data-stu-id="86b7a-182">1.3</span></span>|
|[<span data-ttu-id="86b7a-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-184">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="86b7a-184">Restricted</span></span>|
|[<span data-ttu-id="86b7a-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="86b7a-187">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="86b7a-187">Returns:</span></span>

<span data-ttu-id="86b7a-188">Тип: String</span><span class="sxs-lookup"><span data-stu-id="86b7a-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="86b7a-189">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="86b7a-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="86b7a-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="86b7a-191">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="86b7a-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="86b7a-p105">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p105">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="86b7a-p106">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p106">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-197">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-197">Parameters</span></span>

|<span data-ttu-id="86b7a-198">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-198">Name</span></span>| <span data-ttu-id="86b7a-199">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-199">Type</span></span>| <span data-ttu-id="86b7a-200">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="86b7a-201">Date</span><span class="sxs-lookup"><span data-stu-id="86b7a-201">Date</span></span>|<span data-ttu-id="86b7a-202">Объект Date</span><span class="sxs-lookup"><span data-stu-id="86b7a-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-203">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-203">Requirements</span></span>

|<span data-ttu-id="86b7a-204">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-204">Requirement</span></span>| <span data-ttu-id="86b7a-205">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-206">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-207">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-207">1.0</span></span>|
|[<span data-ttu-id="86b7a-208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-209">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="86b7a-212">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="86b7a-212">Returns:</span></span>

<span data-ttu-id="86b7a-213">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="86b7a-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="86b7a-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="86b7a-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="86b7a-215">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="86b7a-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-216">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="86b7a-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="86b7a-p107">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-219">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-219">Parameters</span></span>

|<span data-ttu-id="86b7a-220">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-220">Name</span></span>| <span data-ttu-id="86b7a-221">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-221">Type</span></span>| <span data-ttu-id="86b7a-222">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="86b7a-223">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-223">String</span></span>|<span data-ttu-id="86b7a-224">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="86b7a-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="86b7a-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="86b7a-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="86b7a-226">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="86b7a-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-227">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-227">Requirements</span></span>

|<span data-ttu-id="86b7a-228">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-228">Requirement</span></span>| <span data-ttu-id="86b7a-229">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-230">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="86b7a-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-231">1.3</span><span class="sxs-lookup"><span data-stu-id="86b7a-231">1.3</span></span>|
|[<span data-ttu-id="86b7a-232">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-233">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="86b7a-233">Restricted</span></span>|
|[<span data-ttu-id="86b7a-234">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-235">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="86b7a-236">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="86b7a-236">Returns:</span></span>

<span data-ttu-id="86b7a-237">Тип: String</span><span class="sxs-lookup"><span data-stu-id="86b7a-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="86b7a-238">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="86b7a-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="86b7a-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="86b7a-240">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="86b7a-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="86b7a-241">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="86b7a-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-242">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-242">Parameters</span></span>

|<span data-ttu-id="86b7a-243">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-243">Name</span></span>| <span data-ttu-id="86b7a-244">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-244">Type</span></span>| <span data-ttu-id="86b7a-245">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="86b7a-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="86b7a-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|<span data-ttu-id="86b7a-247">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="86b7a-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-248">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-248">Requirements</span></span>

|<span data-ttu-id="86b7a-249">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-249">Requirement</span></span>| <span data-ttu-id="86b7a-250">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-252">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-252">1.0</span></span>|
|[<span data-ttu-id="86b7a-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-254">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-256">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="86b7a-257">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="86b7a-257">Returns:</span></span>

<span data-ttu-id="86b7a-258">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="86b7a-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="86b7a-259">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="86b7a-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="86b7a-260">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="86b7a-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="86b7a-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="86b7a-262">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="86b7a-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-263">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="86b7a-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="86b7a-264">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="86b7a-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="86b7a-p108">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="86b7a-p108">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="86b7a-267">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="86b7a-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="86b7a-268">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="86b7a-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-269">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-269">Parameters</span></span>

|<span data-ttu-id="86b7a-270">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-270">Name</span></span>| <span data-ttu-id="86b7a-271">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-271">Type</span></span>| <span data-ttu-id="86b7a-272">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="86b7a-273">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-273">String</span></span>|<span data-ttu-id="86b7a-274">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="86b7a-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-275">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-275">Requirements</span></span>

|<span data-ttu-id="86b7a-276">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-276">Requirement</span></span>| <span data-ttu-id="86b7a-277">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-278">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-279">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-279">1.0</span></span>|
|[<span data-ttu-id="86b7a-280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-281">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-283">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="86b7a-284">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="86b7a-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="86b7a-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="86b7a-286">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="86b7a-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-287">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="86b7a-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="86b7a-288">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="86b7a-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="86b7a-289">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="86b7a-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="86b7a-290">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="86b7a-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="86b7a-p109">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-293">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-293">Parameters</span></span>

|<span data-ttu-id="86b7a-294">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-294">Name</span></span>| <span data-ttu-id="86b7a-295">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-295">Type</span></span>| <span data-ttu-id="86b7a-296">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="86b7a-297">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-297">String</span></span>|<span data-ttu-id="86b7a-298">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="86b7a-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-299">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-299">Requirements</span></span>

|<span data-ttu-id="86b7a-300">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-300">Requirement</span></span>| <span data-ttu-id="86b7a-301">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-302">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-303">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-303">1.0</span></span>|
|[<span data-ttu-id="86b7a-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-305">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-307">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="86b7a-308">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="86b7a-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="86b7a-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="86b7a-310">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="86b7a-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-311">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="86b7a-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="86b7a-p110">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="86b7a-p111">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p111">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="86b7a-p112">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="86b7a-319">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="86b7a-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-320">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-320">Parameters</span></span>

|<span data-ttu-id="86b7a-321">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-321">Name</span></span>| <span data-ttu-id="86b7a-322">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-322">Type</span></span>| <span data-ttu-id="86b7a-323">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="86b7a-324">Object</span><span class="sxs-lookup"><span data-stu-id="86b7a-324">Object</span></span> | <span data-ttu-id="86b7a-325">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="86b7a-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="86b7a-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="86b7a-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="86b7a-p113">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="86b7a-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="86b7a-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="86b7a-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="86b7a-332">Date</span><span class="sxs-lookup"><span data-stu-id="86b7a-332">Date</span></span> | <span data-ttu-id="86b7a-333">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="86b7a-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="86b7a-334">Date</span><span class="sxs-lookup"><span data-stu-id="86b7a-334">Date</span></span> | <span data-ttu-id="86b7a-335">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="86b7a-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="86b7a-336">Строка</span><span class="sxs-lookup"><span data-stu-id="86b7a-336">String</span></span> | <span data-ttu-id="86b7a-p115">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="86b7a-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="86b7a-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="86b7a-p116">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="86b7a-342">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-342">String</span></span> | <span data-ttu-id="86b7a-p117">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="86b7a-345">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-345">String</span></span> | <span data-ttu-id="86b7a-p118">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="86b7a-348">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-348">Requirements</span></span>

|<span data-ttu-id="86b7a-349">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-349">Requirement</span></span>| <span data-ttu-id="86b7a-350">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-351">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-352">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-352">1.0</span></span>|
|[<span data-ttu-id="86b7a-353">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-354">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-355">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-356">Чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="86b7a-357">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="86b7a-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="86b7a-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="86b7a-359">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="86b7a-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="86b7a-p119">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="86b7a-362">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="86b7a-362">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="86b7a-363">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="86b7a-363">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="86b7a-364">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="86b7a-364">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="86b7a-365">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="86b7a-365">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="86b7a-366">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="86b7a-366">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="86b7a-367">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="86b7a-367">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-368">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-368">Parameters</span></span>

|<span data-ttu-id="86b7a-369">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-369">Name</span></span>| <span data-ttu-id="86b7a-370">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-370">Type</span></span>| <span data-ttu-id="86b7a-371">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="86b7a-371">Attributes</span></span>| <span data-ttu-id="86b7a-372">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="86b7a-373">функция</span><span class="sxs-lookup"><span data-stu-id="86b7a-373">function</span></span>||<span data-ttu-id="86b7a-374">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="86b7a-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="86b7a-375">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="86b7a-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="86b7a-376">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="86b7a-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="86b7a-377">Объект</span><span class="sxs-lookup"><span data-stu-id="86b7a-377">Object</span></span>| <span data-ttu-id="86b7a-378">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="86b7a-378">&lt;optional&gt;</span></span>|<span data-ttu-id="86b7a-379">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="86b7a-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="86b7a-380">Ошибки</span><span class="sxs-lookup"><span data-stu-id="86b7a-380">Errors</span></span>

|<span data-ttu-id="86b7a-381">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="86b7a-381">Error code</span></span>|<span data-ttu-id="86b7a-382">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="86b7a-383">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="86b7a-383">The request has failed.</span></span> <span data-ttu-id="86b7a-384">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="86b7a-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="86b7a-385">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="86b7a-385">The Exchange server returned an error.</span></span> <span data-ttu-id="86b7a-386">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="86b7a-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="86b7a-387">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="86b7a-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="86b7a-388">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="86b7a-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-389">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-389">Requirements</span></span>

|<span data-ttu-id="86b7a-390">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="86b7a-391">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-392">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-392">1.0</span></span> | <span data-ttu-id="86b7a-393">1.3</span><span class="sxs-lookup"><span data-stu-id="86b7a-393">1.3</span></span> |
|[<span data-ttu-id="86b7a-394">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-395">ReadItem</span></span> | <span data-ttu-id="86b7a-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-396">ReadItem</span></span> |
|[<span data-ttu-id="86b7a-397">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-398">Чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-398">Read</span></span> | <span data-ttu-id="86b7a-399">Создание</span><span class="sxs-lookup"><span data-stu-id="86b7a-399">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="86b7a-400">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-400">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="86b7a-401">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="86b7a-401">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="86b7a-402">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="86b7a-402">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="86b7a-403">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="86b7a-403">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-404">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-404">Parameters</span></span>

|<span data-ttu-id="86b7a-405">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-405">Name</span></span>| <span data-ttu-id="86b7a-406">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-406">Type</span></span>| <span data-ttu-id="86b7a-407">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="86b7a-407">Attributes</span></span>| <span data-ttu-id="86b7a-408">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-408">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="86b7a-409">функция</span><span class="sxs-lookup"><span data-stu-id="86b7a-409">function</span></span>||<span data-ttu-id="86b7a-410">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="86b7a-410">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="86b7a-411">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="86b7a-411">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="86b7a-412">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="86b7a-412">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="86b7a-413">Объект</span><span class="sxs-lookup"><span data-stu-id="86b7a-413">Object</span></span>| <span data-ttu-id="86b7a-414">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="86b7a-414">&lt;optional&gt;</span></span>|<span data-ttu-id="86b7a-415">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="86b7a-415">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="86b7a-416">Ошибки</span><span class="sxs-lookup"><span data-stu-id="86b7a-416">Errors</span></span>

|<span data-ttu-id="86b7a-417">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="86b7a-417">Error code</span></span>|<span data-ttu-id="86b7a-418">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-418">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="86b7a-419">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="86b7a-419">The request has failed.</span></span> <span data-ttu-id="86b7a-420">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="86b7a-420">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="86b7a-421">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="86b7a-421">The Exchange server returned an error.</span></span> <span data-ttu-id="86b7a-422">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="86b7a-422">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="86b7a-423">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="86b7a-423">The user is no longer connected to the network.</span></span> <span data-ttu-id="86b7a-424">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="86b7a-424">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-425">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-425">Requirements</span></span>

|<span data-ttu-id="86b7a-426">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-426">Requirement</span></span>| <span data-ttu-id="86b7a-427">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-428">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-429">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-429">1.0</span></span>|
|[<span data-ttu-id="86b7a-430">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="86b7a-431">ReadItem</span></span>|
|[<span data-ttu-id="86b7a-432">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-433">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="86b7a-434">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-434">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="86b7a-435">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="86b7a-435">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="86b7a-436">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="86b7a-436">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-437">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="86b7a-437">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="86b7a-438">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="86b7a-438">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="86b7a-439">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="86b7a-439">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="86b7a-440">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="86b7a-440">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="86b7a-441">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="86b7a-441">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="86b7a-442">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="86b7a-442">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="86b7a-443">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="86b7a-443">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="86b7a-444">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="86b7a-444">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="86b7a-p129">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="86b7a-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="86b7a-447">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="86b7a-447">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="86b7a-448">Различия версий</span><span class="sxs-lookup"><span data-stu-id="86b7a-448">Version differences</span></span>

<span data-ttu-id="86b7a-449">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="86b7a-449">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="86b7a-p130">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="86b7a-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="86b7a-453">Параметры</span><span class="sxs-lookup"><span data-stu-id="86b7a-453">Parameters</span></span>

|<span data-ttu-id="86b7a-454">Имя</span><span class="sxs-lookup"><span data-stu-id="86b7a-454">Name</span></span>| <span data-ttu-id="86b7a-455">Тип</span><span class="sxs-lookup"><span data-stu-id="86b7a-455">Type</span></span>| <span data-ttu-id="86b7a-456">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="86b7a-456">Attributes</span></span>| <span data-ttu-id="86b7a-457">Описание</span><span class="sxs-lookup"><span data-stu-id="86b7a-457">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="86b7a-458">String</span><span class="sxs-lookup"><span data-stu-id="86b7a-458">String</span></span>||<span data-ttu-id="86b7a-459">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="86b7a-459">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="86b7a-460">function</span><span class="sxs-lookup"><span data-stu-id="86b7a-460">function</span></span>||<span data-ttu-id="86b7a-461">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="86b7a-461">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="86b7a-462">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="86b7a-462">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="86b7a-463">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="86b7a-463">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="86b7a-464">Object</span><span class="sxs-lookup"><span data-stu-id="86b7a-464">Object</span></span>| <span data-ttu-id="86b7a-465">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="86b7a-465">&lt;optional&gt;</span></span>|<span data-ttu-id="86b7a-466">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="86b7a-466">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86b7a-467">Требования</span><span class="sxs-lookup"><span data-stu-id="86b7a-467">Requirements</span></span>

|<span data-ttu-id="86b7a-468">Требование</span><span class="sxs-lookup"><span data-stu-id="86b7a-468">Requirement</span></span>| <span data-ttu-id="86b7a-469">Значение</span><span class="sxs-lookup"><span data-stu-id="86b7a-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="86b7a-470">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="86b7a-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86b7a-471">1.0</span><span class="sxs-lookup"><span data-stu-id="86b7a-471">1.0</span></span>|
|[<span data-ttu-id="86b7a-472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="86b7a-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="86b7a-473">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="86b7a-473">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="86b7a-474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="86b7a-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="86b7a-475">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="86b7a-475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="86b7a-476">Пример</span><span class="sxs-lookup"><span data-stu-id="86b7a-476">Example</span></span>

<span data-ttu-id="86b7a-477">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="86b7a-477">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
