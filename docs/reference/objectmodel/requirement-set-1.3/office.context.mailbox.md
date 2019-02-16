---
title: Office.context.mailbox — набор обязательных элементов 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: edf9bcaa842660007bea85b497308aa7949f8085
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068226"
---
# <a name="mailbox"></a><span data-ttu-id="1a1b4-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="1a1b4-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="1a1b4-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="1a1b4-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="1a1b4-104">Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a1b4-105">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-105">Requirements</span></span>

|<span data-ttu-id="1a1b4-106">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-106">Requirement</span></span>| <span data-ttu-id="1a1b4-107">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-109">1.0</span></span>|
|[<span data-ttu-id="1a1b4-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1a1b4-111">Restricted</span></span>|
|[<span data-ttu-id="1a1b4-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="1a1b4-114">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="1a1b4-114">Namespaces</span></span>

<span data-ttu-id="1a1b4-115">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="1a1b4-116">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="1a1b4-117">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="1a1b4-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="1a1b4-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="1a1b4-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-119">ewsUrl :String</span></span>

<span data-ttu-id="1a1b4-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-122">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a1b4-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1a1b4-125">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="1a1b4-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1a1b4-128">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-128">Type</span></span>

*   <span data-ttu-id="1a1b4-129">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a1b4-130">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-130">Requirements</span></span>

|<span data-ttu-id="1a1b4-131">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-131">Requirement</span></span>| <span data-ttu-id="1a1b4-132">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-134">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-134">1.0</span></span>|
|[<span data-ttu-id="1a1b4-135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-136">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1a1b4-139">Методы</span><span class="sxs-lookup"><span data-stu-id="1a1b4-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="1a1b4-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1a1b4-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1a1b4-141">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-142">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a1b4-p104">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-145">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-145">Parameters</span></span>

|<span data-ttu-id="1a1b4-146">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-146">Name</span></span>| <span data-ttu-id="1a1b4-147">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-147">Type</span></span>| <span data-ttu-id="1a1b4-148">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1a1b4-149">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-149">String</span></span>|<span data-ttu-id="1a1b4-150">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="1a1b4-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1a1b4-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="1a1b4-152">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-153">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-153">Requirements</span></span>

|<span data-ttu-id="1a1b4-154">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-154">Requirement</span></span>| <span data-ttu-id="1a1b4-155">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-156">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1a1b4-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-157">1.3</span><span class="sxs-lookup"><span data-stu-id="1a1b4-157">1.3</span></span>|
|[<span data-ttu-id="1a1b4-158">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-159">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1a1b4-159">Restricted</span></span>|
|[<span data-ttu-id="1a1b4-160">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-161">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a1b4-162">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1a1b4-162">Returns:</span></span>

<span data-ttu-id="1a1b4-163">Тип: String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1a1b4-164">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="1a1b4-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="1a1b4-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="1a1b4-166">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="1a1b4-p105">В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p105">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="1a1b4-p106">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p106">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-172">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-172">Parameters</span></span>

|<span data-ttu-id="1a1b4-173">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-173">Name</span></span>| <span data-ttu-id="1a1b4-174">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-174">Type</span></span>| <span data-ttu-id="1a1b4-175">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="1a1b4-176">Date</span><span class="sxs-lookup"><span data-stu-id="1a1b4-176">Date</span></span>|<span data-ttu-id="1a1b4-177">Объект Date</span><span class="sxs-lookup"><span data-stu-id="1a1b4-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-178">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-178">Requirements</span></span>

|<span data-ttu-id="1a1b4-179">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-179">Requirement</span></span>| <span data-ttu-id="1a1b4-180">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-182">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-182">1.0</span></span>|
|[<span data-ttu-id="1a1b4-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-184">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a1b4-187">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1a1b4-187">Returns:</span></span>

<span data-ttu-id="1a1b4-188">Тип: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="1a1b4-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="1a1b4-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1a1b4-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1a1b4-190">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-191">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a1b4-p107">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-194">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-194">Parameters</span></span>

|<span data-ttu-id="1a1b4-195">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-195">Name</span></span>| <span data-ttu-id="1a1b4-196">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-196">Type</span></span>| <span data-ttu-id="1a1b4-197">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1a1b4-198">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-198">String</span></span>|<span data-ttu-id="1a1b4-199">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="1a1b4-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="1a1b4-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1a1b4-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="1a1b4-201">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-202">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-202">Requirements</span></span>

|<span data-ttu-id="1a1b4-203">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-203">Requirement</span></span>| <span data-ttu-id="1a1b4-204">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-205">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1a1b4-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-206">1.3</span><span class="sxs-lookup"><span data-stu-id="1a1b4-206">1.3</span></span>|
|[<span data-ttu-id="1a1b4-207">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-208">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1a1b4-208">Restricted</span></span>|
|[<span data-ttu-id="1a1b4-209">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-210">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a1b4-211">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1a1b4-211">Returns:</span></span>

<span data-ttu-id="1a1b4-212">Тип: String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1a1b4-213">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="1a1b4-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="1a1b4-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="1a1b4-215">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="1a1b4-216">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-217">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-217">Parameters</span></span>

|<span data-ttu-id="1a1b4-218">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-218">Name</span></span>| <span data-ttu-id="1a1b4-219">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-219">Type</span></span>| <span data-ttu-id="1a1b4-220">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="1a1b4-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1a1b4-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="1a1b4-222">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-223">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-223">Requirements</span></span>

|<span data-ttu-id="1a1b4-224">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-224">Requirement</span></span>| <span data-ttu-id="1a1b4-225">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-226">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-227">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-227">1.0</span></span>|
|[<span data-ttu-id="1a1b4-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-229">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-231">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a1b4-232">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1a1b4-232">Returns:</span></span>

<span data-ttu-id="1a1b4-233">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="1a1b4-234">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1a1b4-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1a1b4-235">Date</span><span class="sxs-lookup"><span data-stu-id="1a1b4-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="1a1b4-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1a1b4-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="1a1b4-237">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-238">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a1b4-239">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1a1b4-p108">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p108">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="1a1b4-242">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="1a1b4-243">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-244">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-244">Parameters</span></span>

|<span data-ttu-id="1a1b4-245">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-245">Name</span></span>| <span data-ttu-id="1a1b4-246">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-246">Type</span></span>| <span data-ttu-id="1a1b4-247">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1a1b4-248">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-248">String</span></span>|<span data-ttu-id="1a1b4-249">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-250">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-250">Requirements</span></span>

|<span data-ttu-id="1a1b4-251">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-251">Requirement</span></span>| <span data-ttu-id="1a1b4-252">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-253">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-254">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-254">1.0</span></span>|
|[<span data-ttu-id="1a1b4-255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-256">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-258">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a1b4-259">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="1a1b4-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1a1b4-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="1a1b4-261">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-262">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a1b4-263">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1a1b4-264">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="1a1b4-265">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="1a1b4-p109">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-268">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-268">Parameters</span></span>

|<span data-ttu-id="1a1b4-269">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-269">Name</span></span>| <span data-ttu-id="1a1b4-270">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-270">Type</span></span>| <span data-ttu-id="1a1b4-271">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1a1b4-272">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-272">String</span></span>|<span data-ttu-id="1a1b4-273">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-274">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-274">Requirements</span></span>

|<span data-ttu-id="1a1b4-275">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-275">Requirement</span></span>| <span data-ttu-id="1a1b4-276">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-277">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-278">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-278">1.0</span></span>|
|[<span data-ttu-id="1a1b4-279">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-280">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-281">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-282">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a1b4-283">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="1a1b4-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1a1b4-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="1a1b4-285">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-286">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a1b4-p110">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1a1b4-p111">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p111">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="1a1b4-p112">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="1a1b4-294">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-295">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-295">Parameters</span></span>

|<span data-ttu-id="1a1b4-296">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-296">Name</span></span>| <span data-ttu-id="1a1b4-297">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-297">Type</span></span>| <span data-ttu-id="1a1b4-298">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1a1b4-299">Object</span><span class="sxs-lookup"><span data-stu-id="1a1b4-299">Object</span></span> | <span data-ttu-id="1a1b4-300">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="1a1b4-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1a1b4-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1a1b4-p113">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="1a1b4-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1a1b4-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1a1b4-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="1a1b4-307">Date</span><span class="sxs-lookup"><span data-stu-id="1a1b4-307">Date</span></span> | <span data-ttu-id="1a1b4-308">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="1a1b4-309">Date</span><span class="sxs-lookup"><span data-stu-id="1a1b4-309">Date</span></span> | <span data-ttu-id="1a1b4-310">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="1a1b4-311">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-311">String</span></span> | <span data-ttu-id="1a1b4-p115">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="1a1b4-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="1a1b4-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="1a1b4-p116">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1a1b4-317">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-317">String</span></span> | <span data-ttu-id="1a1b4-p117">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="1a1b4-320">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-320">String</span></span> | <span data-ttu-id="1a1b4-p118">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a1b4-323">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-323">Requirements</span></span>

|<span data-ttu-id="1a1b4-324">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-324">Requirement</span></span>| <span data-ttu-id="1a1b4-325">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-326">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-327">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-327">1.0</span></span>|
|[<span data-ttu-id="1a1b4-328">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-329">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-330">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-331">Чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a1b4-332">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="1a1b4-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1a1b4-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1a1b4-334">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="1a1b4-p119">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="1a1b4-p120">Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1a1b4-340">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="1a1b4-p121">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-343">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-343">Parameters</span></span>

|<span data-ttu-id="1a1b4-344">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-344">Name</span></span>| <span data-ttu-id="1a1b4-345">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-345">Type</span></span>| <span data-ttu-id="1a1b4-346">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1a1b4-346">Attributes</span></span>| <span data-ttu-id="1a1b4-347">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1a1b4-348">function</span><span class="sxs-lookup"><span data-stu-id="1a1b4-348">function</span></span>||<span data-ttu-id="1a1b4-p122">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1a1b4-351">Объект</span><span class="sxs-lookup"><span data-stu-id="1a1b4-351">Object</span></span>| <span data-ttu-id="1a1b4-352">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1a1b4-352">&lt;optional&gt;</span></span>|<span data-ttu-id="1a1b4-353">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-354">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-354">Requirements</span></span>

|<span data-ttu-id="1a1b4-355">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-355">Requirement</span></span>| <span data-ttu-id="1a1b4-356">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-357">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1a1b4-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-358">1.3</span><span class="sxs-lookup"><span data-stu-id="1a1b4-358">1.3</span></span>|
|[<span data-ttu-id="1a1b4-359">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-360">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-361">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-362">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a1b4-363">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="1a1b4-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1a1b4-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1a1b4-365">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="1a1b4-366">Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-367">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-367">Parameters</span></span>

|<span data-ttu-id="1a1b4-368">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-368">Name</span></span>| <span data-ttu-id="1a1b4-369">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-369">Type</span></span>| <span data-ttu-id="1a1b4-370">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1a1b4-370">Attributes</span></span>| <span data-ttu-id="1a1b4-371">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1a1b4-372">функция</span><span class="sxs-lookup"><span data-stu-id="1a1b4-372">function</span></span>||<span data-ttu-id="1a1b4-373">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1a1b4-374">Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1a1b4-375">Object</span><span class="sxs-lookup"><span data-stu-id="1a1b4-375">Object</span></span>| <span data-ttu-id="1a1b4-376">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1a1b4-376">&lt;optional&gt;</span></span>|<span data-ttu-id="1a1b4-377">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-378">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-378">Requirements</span></span>

|<span data-ttu-id="1a1b4-379">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-379">Requirement</span></span>| <span data-ttu-id="1a1b4-380">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-381">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-382">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-382">1.0</span></span>|
|[<span data-ttu-id="1a1b4-383">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a1b4-384">ReadItem</span></span>|
|[<span data-ttu-id="1a1b4-385">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-386">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a1b4-387">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="1a1b4-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1a1b4-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="1a1b4-389">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-390">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="1a1b4-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="1a1b4-391">В Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="1a1b4-392">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="1a1b4-393">В таких случаях надстройка должна [использовать REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="1a1b4-394">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="1a1b4-395">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="1a1b4-396">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="1a1b4-397">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="1a1b4-p124">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="1a1b4-400">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="1a1b4-401">Различия версий</span><span class="sxs-lookup"><span data-stu-id="1a1b4-401">Version differences</span></span>

<span data-ttu-id="1a1b4-402">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="1a1b4-p125">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a1b4-406">Параметры</span><span class="sxs-lookup"><span data-stu-id="1a1b4-406">Parameters</span></span>

|<span data-ttu-id="1a1b4-407">Имя</span><span class="sxs-lookup"><span data-stu-id="1a1b4-407">Name</span></span>| <span data-ttu-id="1a1b4-408">Тип</span><span class="sxs-lookup"><span data-stu-id="1a1b4-408">Type</span></span>| <span data-ttu-id="1a1b4-409">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1a1b4-409">Attributes</span></span>| <span data-ttu-id="1a1b4-410">Описание</span><span class="sxs-lookup"><span data-stu-id="1a1b4-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1a1b4-411">String</span><span class="sxs-lookup"><span data-stu-id="1a1b4-411">String</span></span>||<span data-ttu-id="1a1b4-412">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="1a1b4-413">function</span><span class="sxs-lookup"><span data-stu-id="1a1b4-413">function</span></span>||<span data-ttu-id="1a1b4-414">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1a1b4-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1a1b4-415">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="1a1b4-416">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="1a1b4-417">Объект</span><span class="sxs-lookup"><span data-stu-id="1a1b4-417">Object</span></span>| <span data-ttu-id="1a1b4-418">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1a1b4-418">&lt;optional&gt;</span></span>|<span data-ttu-id="1a1b4-419">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a1b4-420">Требования</span><span class="sxs-lookup"><span data-stu-id="1a1b4-420">Requirements</span></span>

|<span data-ttu-id="1a1b4-421">Требование</span><span class="sxs-lookup"><span data-stu-id="1a1b4-421">Requirement</span></span>| <span data-ttu-id="1a1b4-422">Значение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a1b4-423">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1a1b4-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a1b4-424">1.0</span><span class="sxs-lookup"><span data-stu-id="1a1b4-424">1.0</span></span>|
|[<span data-ttu-id="1a1b4-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1a1b4-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a1b4-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1a1b4-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="1a1b4-427">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1a1b4-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1a1b4-428">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1a1b4-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a1b4-429">Пример</span><span class="sxs-lookup"><span data-stu-id="1a1b4-429">Example</span></span>

<span data-ttu-id="1a1b4-430">В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="1a1b4-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
