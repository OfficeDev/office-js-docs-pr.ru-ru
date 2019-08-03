---
title: Office. Context. Mailbox — набор обязательных элементов 1,4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 373ac1f4361ae94f788a61e49c304a1cfdb90232
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064706"
---
# <a name="mailbox"></a><span data-ttu-id="37b7a-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="37b7a-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="37b7a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="37b7a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="37b7a-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="37b7a-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37b7a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-105">Requirements</span></span>

|<span data-ttu-id="37b7a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-106">Requirement</span></span>| <span data-ttu-id="37b7a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-109">1.0</span></span>|
|[<span data-ttu-id="37b7a-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="37b7a-111">Restricted</span></span>|
|[<span data-ttu-id="37b7a-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="37b7a-114">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="37b7a-114">Namespaces</span></span>

<span data-ttu-id="37b7a-115">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="37b7a-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="37b7a-116">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="37b7a-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="37b7a-117">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="37b7a-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="37b7a-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="37b7a-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="37b7a-119">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="37b7a-119">ewsUrl: String</span></span>

<span data-ttu-id="37b7a-120">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="37b7a-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="37b7a-121">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37b7a-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-122">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="37b7a-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="37b7a-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="37b7a-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="37b7a-125">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="37b7a-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="37b7a-128">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-128">Type</span></span>

*   <span data-ttu-id="37b7a-129">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37b7a-130">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-130">Requirements</span></span>

|<span data-ttu-id="37b7a-131">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-131">Requirement</span></span>| <span data-ttu-id="37b7a-132">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-134">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-134">1.0</span></span>|
|[<span data-ttu-id="37b7a-135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-136">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="37b7a-139">Методы</span><span class="sxs-lookup"><span data-stu-id="37b7a-139">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="37b7a-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="37b7a-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="37b7a-141">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="37b7a-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-142">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="37b7a-142">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="37b7a-p104">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-145">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-145">Parameters</span></span>

|<span data-ttu-id="37b7a-146">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-146">Name</span></span>| <span data-ttu-id="37b7a-147">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-147">Type</span></span>| <span data-ttu-id="37b7a-148">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="37b7a-149">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-149">String</span></span>|<span data-ttu-id="37b7a-150">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="37b7a-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="37b7a-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="37b7a-152">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="37b7a-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-153">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-153">Requirements</span></span>

|<span data-ttu-id="37b7a-154">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-154">Requirement</span></span>| <span data-ttu-id="37b7a-155">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-156">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37b7a-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-157">1.3</span><span class="sxs-lookup"><span data-stu-id="37b7a-157">1.3</span></span>|
|[<span data-ttu-id="37b7a-158">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-158">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-159">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="37b7a-159">Restricted</span></span>|
|[<span data-ttu-id="37b7a-160">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-160">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-161">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37b7a-162">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37b7a-162">Returns:</span></span>

<span data-ttu-id="37b7a-163">Тип: String</span><span class="sxs-lookup"><span data-stu-id="37b7a-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="37b7a-164">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="37b7a-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="37b7a-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="37b7a-166">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="37b7a-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="37b7a-167">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="37b7a-167">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="37b7a-168">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="37b7a-168">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="37b7a-169">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="37b7a-169">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="37b7a-170">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="37b7a-170">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="37b7a-171">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="37b7a-171">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-172">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-172">Parameters</span></span>

|<span data-ttu-id="37b7a-173">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-173">Name</span></span>| <span data-ttu-id="37b7a-174">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-174">Type</span></span>| <span data-ttu-id="37b7a-175">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="37b7a-176">Date</span><span class="sxs-lookup"><span data-stu-id="37b7a-176">Date</span></span>|<span data-ttu-id="37b7a-177">Объект Date</span><span class="sxs-lookup"><span data-stu-id="37b7a-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-178">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-178">Requirements</span></span>

|<span data-ttu-id="37b7a-179">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-179">Requirement</span></span>| <span data-ttu-id="37b7a-180">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-182">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-182">1.0</span></span>|
|[<span data-ttu-id="37b7a-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-184">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37b7a-187">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37b7a-187">Returns:</span></span>

<span data-ttu-id="37b7a-188">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="37b7a-188">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="37b7a-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="37b7a-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="37b7a-190">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="37b7a-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-191">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="37b7a-191">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="37b7a-p107">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-194">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-194">Parameters</span></span>

|<span data-ttu-id="37b7a-195">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-195">Name</span></span>| <span data-ttu-id="37b7a-196">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-196">Type</span></span>| <span data-ttu-id="37b7a-197">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="37b7a-198">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-198">String</span></span>|<span data-ttu-id="37b7a-199">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="37b7a-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="37b7a-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="37b7a-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="37b7a-201">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="37b7a-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-202">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-202">Requirements</span></span>

|<span data-ttu-id="37b7a-203">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-203">Requirement</span></span>| <span data-ttu-id="37b7a-204">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-205">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37b7a-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-206">1.3</span><span class="sxs-lookup"><span data-stu-id="37b7a-206">1.3</span></span>|
|[<span data-ttu-id="37b7a-207">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-207">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-208">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="37b7a-208">Restricted</span></span>|
|[<span data-ttu-id="37b7a-209">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-209">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-210">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37b7a-211">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37b7a-211">Returns:</span></span>

<span data-ttu-id="37b7a-212">Тип: String</span><span class="sxs-lookup"><span data-stu-id="37b7a-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="37b7a-213">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="37b7a-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="37b7a-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="37b7a-215">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="37b7a-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="37b7a-216">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="37b7a-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-217">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-217">Parameters</span></span>

|<span data-ttu-id="37b7a-218">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-218">Name</span></span>| <span data-ttu-id="37b7a-219">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-219">Type</span></span>| <span data-ttu-id="37b7a-220">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="37b7a-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="37b7a-221">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="37b7a-222">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="37b7a-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-223">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-223">Requirements</span></span>

|<span data-ttu-id="37b7a-224">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-224">Requirement</span></span>| <span data-ttu-id="37b7a-225">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-226">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-227">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-227">1.0</span></span>|
|[<span data-ttu-id="37b7a-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-228">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-229">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-230">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-231">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37b7a-232">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37b7a-232">Returns:</span></span>

<span data-ttu-id="37b7a-233">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="37b7a-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="37b7a-234">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="37b7a-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37b7a-235">Дата</span><span class="sxs-lookup"><span data-stu-id="37b7a-235">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="37b7a-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="37b7a-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="37b7a-237">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="37b7a-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-238">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="37b7a-238">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="37b7a-239">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="37b7a-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="37b7a-240">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="37b7a-240">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="37b7a-241">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="37b7a-241">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="37b7a-242">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="37b7a-242">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="37b7a-243">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="37b7a-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-244">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-244">Parameters</span></span>

|<span data-ttu-id="37b7a-245">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-245">Name</span></span>| <span data-ttu-id="37b7a-246">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-246">Type</span></span>| <span data-ttu-id="37b7a-247">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="37b7a-248">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-248">String</span></span>|<span data-ttu-id="37b7a-249">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="37b7a-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-250">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-250">Requirements</span></span>

|<span data-ttu-id="37b7a-251">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-251">Requirement</span></span>| <span data-ttu-id="37b7a-252">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-253">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-254">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-254">1.0</span></span>|
|[<span data-ttu-id="37b7a-255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-256">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-258">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37b7a-259">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="37b7a-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="37b7a-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="37b7a-261">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="37b7a-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-262">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="37b7a-262">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="37b7a-263">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="37b7a-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="37b7a-264">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="37b7a-264">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="37b7a-265">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="37b7a-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="37b7a-p109">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-268">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-268">Parameters</span></span>

|<span data-ttu-id="37b7a-269">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-269">Name</span></span>| <span data-ttu-id="37b7a-270">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-270">Type</span></span>| <span data-ttu-id="37b7a-271">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="37b7a-272">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-272">String</span></span>|<span data-ttu-id="37b7a-273">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="37b7a-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-274">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-274">Requirements</span></span>

|<span data-ttu-id="37b7a-275">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-275">Requirement</span></span>| <span data-ttu-id="37b7a-276">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-277">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-278">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-278">1.0</span></span>|
|[<span data-ttu-id="37b7a-279">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-280">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-281">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-282">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37b7a-283">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="37b7a-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="37b7a-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="37b7a-285">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="37b7a-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-286">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="37b7a-286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="37b7a-p110">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="37b7a-289">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="37b7a-289">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="37b7a-290">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-290">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="37b7a-291">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-291">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="37b7a-p112">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="37b7a-294">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="37b7a-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-295">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-295">Parameters</span></span>

|<span data-ttu-id="37b7a-296">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-296">Name</span></span>| <span data-ttu-id="37b7a-297">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-297">Type</span></span>| <span data-ttu-id="37b7a-298">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="37b7a-299">Object</span><span class="sxs-lookup"><span data-stu-id="37b7a-299">Object</span></span> | <span data-ttu-id="37b7a-300">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="37b7a-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="37b7a-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="37b7a-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="37b7a-p113">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="37b7a-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="37b7a-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="37b7a-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="37b7a-307">Date</span><span class="sxs-lookup"><span data-stu-id="37b7a-307">Date</span></span> | <span data-ttu-id="37b7a-308">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="37b7a-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="37b7a-309">Date</span><span class="sxs-lookup"><span data-stu-id="37b7a-309">Date</span></span> | <span data-ttu-id="37b7a-310">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="37b7a-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="37b7a-311">Строка</span><span class="sxs-lookup"><span data-stu-id="37b7a-311">String</span></span> | <span data-ttu-id="37b7a-p115">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="37b7a-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="37b7a-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="37b7a-p116">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="37b7a-317">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-317">String</span></span> | <span data-ttu-id="37b7a-p117">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="37b7a-320">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-320">String</span></span> | <span data-ttu-id="37b7a-p118">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37b7a-323">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-323">Requirements</span></span>

|<span data-ttu-id="37b7a-324">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-324">Requirement</span></span>| <span data-ttu-id="37b7a-325">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-326">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-327">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-327">1.0</span></span>|
|[<span data-ttu-id="37b7a-328">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-329">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-330">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-331">Чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37b7a-332">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="37b7a-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="37b7a-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="37b7a-334">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="37b7a-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="37b7a-p119">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="37b7a-p120">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="37b7a-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="37b7a-340">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="37b7a-p121">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-343">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-343">Parameters</span></span>

|<span data-ttu-id="37b7a-344">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-344">Name</span></span>| <span data-ttu-id="37b7a-345">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-345">Type</span></span>| <span data-ttu-id="37b7a-346">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37b7a-346">Attributes</span></span>| <span data-ttu-id="37b7a-347">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="37b7a-348">function</span><span class="sxs-lookup"><span data-stu-id="37b7a-348">function</span></span>||<span data-ttu-id="37b7a-p122">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="37b7a-351">Объект</span><span class="sxs-lookup"><span data-stu-id="37b7a-351">Object</span></span>| <span data-ttu-id="37b7a-352">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37b7a-352">&lt;optional&gt;</span></span>|<span data-ttu-id="37b7a-353">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="37b7a-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-354">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-354">Requirements</span></span>

|<span data-ttu-id="37b7a-355">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-355">Requirement</span></span>| <span data-ttu-id="37b7a-356">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-357">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37b7a-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-358">1.3</span><span class="sxs-lookup"><span data-stu-id="37b7a-358">1.3</span></span>|
|[<span data-ttu-id="37b7a-359">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-360">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-361">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-362">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="37b7a-363">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="37b7a-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="37b7a-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="37b7a-365">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="37b7a-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="37b7a-366">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="37b7a-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-367">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-367">Parameters</span></span>

|<span data-ttu-id="37b7a-368">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-368">Name</span></span>| <span data-ttu-id="37b7a-369">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-369">Type</span></span>| <span data-ttu-id="37b7a-370">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37b7a-370">Attributes</span></span>| <span data-ttu-id="37b7a-371">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="37b7a-372">функция</span><span class="sxs-lookup"><span data-stu-id="37b7a-372">function</span></span>||<span data-ttu-id="37b7a-373">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37b7a-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37b7a-374">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37b7a-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="37b7a-375">Object</span><span class="sxs-lookup"><span data-stu-id="37b7a-375">Object</span></span>| <span data-ttu-id="37b7a-376">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37b7a-376">&lt;optional&gt;</span></span>|<span data-ttu-id="37b7a-377">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="37b7a-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-378">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-378">Requirements</span></span>

|<span data-ttu-id="37b7a-379">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-379">Requirement</span></span>| <span data-ttu-id="37b7a-380">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-381">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-382">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-382">1.0</span></span>|
|[<span data-ttu-id="37b7a-383">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37b7a-384">ReadItem</span></span>|
|[<span data-ttu-id="37b7a-385">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-386">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37b7a-387">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="37b7a-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="37b7a-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="37b7a-389">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="37b7a-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-390">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="37b7a-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="37b7a-391">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="37b7a-391">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="37b7a-392">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="37b7a-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="37b7a-393">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="37b7a-393">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="37b7a-394">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="37b7a-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="37b7a-395">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="37b7a-395">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="37b7a-396">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="37b7a-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="37b7a-397">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="37b7a-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="37b7a-p124">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="37b7a-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="37b7a-400">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="37b7a-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="37b7a-401">Различия версий</span><span class="sxs-lookup"><span data-stu-id="37b7a-401">Version differences</span></span>

<span data-ttu-id="37b7a-402">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="37b7a-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="37b7a-p125">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="37b7a-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37b7a-406">Параметры</span><span class="sxs-lookup"><span data-stu-id="37b7a-406">Parameters</span></span>

|<span data-ttu-id="37b7a-407">Имя</span><span class="sxs-lookup"><span data-stu-id="37b7a-407">Name</span></span>| <span data-ttu-id="37b7a-408">Тип</span><span class="sxs-lookup"><span data-stu-id="37b7a-408">Type</span></span>| <span data-ttu-id="37b7a-409">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37b7a-409">Attributes</span></span>| <span data-ttu-id="37b7a-410">Описание</span><span class="sxs-lookup"><span data-stu-id="37b7a-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="37b7a-411">String</span><span class="sxs-lookup"><span data-stu-id="37b7a-411">String</span></span>||<span data-ttu-id="37b7a-412">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="37b7a-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="37b7a-413">функция</span><span class="sxs-lookup"><span data-stu-id="37b7a-413">function</span></span>||<span data-ttu-id="37b7a-414">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37b7a-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37b7a-415">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37b7a-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="37b7a-416">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="37b7a-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="37b7a-417">Объект</span><span class="sxs-lookup"><span data-stu-id="37b7a-417">Object</span></span>| <span data-ttu-id="37b7a-418">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37b7a-418">&lt;optional&gt;</span></span>|<span data-ttu-id="37b7a-419">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="37b7a-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37b7a-420">Требования</span><span class="sxs-lookup"><span data-stu-id="37b7a-420">Requirements</span></span>

|<span data-ttu-id="37b7a-421">Требование</span><span class="sxs-lookup"><span data-stu-id="37b7a-421">Requirement</span></span>| <span data-ttu-id="37b7a-422">Значение</span><span class="sxs-lookup"><span data-stu-id="37b7a-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="37b7a-423">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37b7a-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37b7a-424">1.0</span><span class="sxs-lookup"><span data-stu-id="37b7a-424">1.0</span></span>|
|[<span data-ttu-id="37b7a-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37b7a-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37b7a-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="37b7a-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="37b7a-427">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37b7a-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37b7a-428">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37b7a-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37b7a-429">Пример</span><span class="sxs-lookup"><span data-stu-id="37b7a-429">Example</span></span>

<span data-ttu-id="37b7a-430">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="37b7a-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
