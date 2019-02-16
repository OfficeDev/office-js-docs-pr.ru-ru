---
title: Пространство имен Office — набор обязательных элементов 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: ad08895719d809436216d2f0bb455260dbca3b1e
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067890"
---
# <a name="office"></a><span data-ttu-id="9b68d-102">Office</span><span class="sxs-lookup"><span data-stu-id="9b68d-102">Office</span></span>

<span data-ttu-id="9b68d-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9b68d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b68d-105">Требования</span><span class="sxs-lookup"><span data-stu-id="9b68d-105">Requirements</span></span>

|<span data-ttu-id="9b68d-106">Требование</span><span class="sxs-lookup"><span data-stu-id="9b68d-106">Requirement</span></span>| <span data-ttu-id="9b68d-107">Значение</span><span class="sxs-lookup"><span data-stu-id="9b68d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b68d-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9b68d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b68d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9b68d-109">1.0</span></span>|
|[<span data-ttu-id="9b68d-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9b68d-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b68d-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9b68d-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="9b68d-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="9b68d-112">Namespaces</span></span>

<span data-ttu-id="9b68d-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="9b68d-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9b68d-114">[MailboxEnums.](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9b68d-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9b68d-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="9b68d-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9b68d-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9b68d-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="9b68d-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="9b68d-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9b68d-118">Тип</span><span class="sxs-lookup"><span data-stu-id="9b68d-118">Type</span></span>

*   <span data-ttu-id="9b68d-119">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9b68d-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9b68d-120">Properties:</span></span>

|<span data-ttu-id="9b68d-121">Имя</span><span class="sxs-lookup"><span data-stu-id="9b68d-121">Name</span></span>| <span data-ttu-id="9b68d-122">Тип</span><span class="sxs-lookup"><span data-stu-id="9b68d-122">Type</span></span>| <span data-ttu-id="9b68d-123">Описание</span><span class="sxs-lookup"><span data-stu-id="9b68d-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9b68d-124">Для указания</span><span class="sxs-lookup"><span data-stu-id="9b68d-124">String</span></span>|<span data-ttu-id="9b68d-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="9b68d-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9b68d-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="9b68d-126">String</span></span>|<span data-ttu-id="9b68d-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="9b68d-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b68d-128">Требования</span><span class="sxs-lookup"><span data-stu-id="9b68d-128">Requirements</span></span>

|<span data-ttu-id="9b68d-129">Требование</span><span class="sxs-lookup"><span data-stu-id="9b68d-129">Requirement</span></span>| <span data-ttu-id="9b68d-130">Значение</span><span class="sxs-lookup"><span data-stu-id="9b68d-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b68d-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9b68d-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b68d-132">1.0</span><span class="sxs-lookup"><span data-stu-id="9b68d-132">1.0</span></span>|
|[<span data-ttu-id="9b68d-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9b68d-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b68d-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9b68d-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="9b68d-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9b68d-135">CoercionType :String</span></span>

<span data-ttu-id="9b68d-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9b68d-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9b68d-137">Тип</span><span class="sxs-lookup"><span data-stu-id="9b68d-137">Type</span></span>

*   <span data-ttu-id="9b68d-138">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9b68d-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9b68d-139">Properties:</span></span>

|<span data-ttu-id="9b68d-140">Имя</span><span class="sxs-lookup"><span data-stu-id="9b68d-140">Name</span></span>| <span data-ttu-id="9b68d-141">Тип</span><span class="sxs-lookup"><span data-stu-id="9b68d-141">Type</span></span>| <span data-ttu-id="9b68d-142">Описание</span><span class="sxs-lookup"><span data-stu-id="9b68d-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9b68d-143">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-143">String</span></span>|<span data-ttu-id="9b68d-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="9b68d-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9b68d-145">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-145">String</span></span>|<span data-ttu-id="9b68d-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="9b68d-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b68d-147">Требования</span><span class="sxs-lookup"><span data-stu-id="9b68d-147">Requirements</span></span>

|<span data-ttu-id="9b68d-148">Требование</span><span class="sxs-lookup"><span data-stu-id="9b68d-148">Requirement</span></span>| <span data-ttu-id="9b68d-149">Значение</span><span class="sxs-lookup"><span data-stu-id="9b68d-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b68d-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9b68d-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b68d-151">1.0</span><span class="sxs-lookup"><span data-stu-id="9b68d-151">1.0</span></span>|
|[<span data-ttu-id="9b68d-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9b68d-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b68d-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9b68d-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="9b68d-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9b68d-154">SourceProperty :String</span></span>

<span data-ttu-id="9b68d-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9b68d-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9b68d-156">Тип</span><span class="sxs-lookup"><span data-stu-id="9b68d-156">Type</span></span>

*   <span data-ttu-id="9b68d-157">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9b68d-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9b68d-158">Properties:</span></span>

|<span data-ttu-id="9b68d-159">Имя</span><span class="sxs-lookup"><span data-stu-id="9b68d-159">Name</span></span>| <span data-ttu-id="9b68d-160">Тип</span><span class="sxs-lookup"><span data-stu-id="9b68d-160">Type</span></span>| <span data-ttu-id="9b68d-161">Описание</span><span class="sxs-lookup"><span data-stu-id="9b68d-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9b68d-162">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-162">String</span></span>|<span data-ttu-id="9b68d-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="9b68d-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9b68d-164">String</span><span class="sxs-lookup"><span data-stu-id="9b68d-164">String</span></span>|<span data-ttu-id="9b68d-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="9b68d-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b68d-166">Требования</span><span class="sxs-lookup"><span data-stu-id="9b68d-166">Requirements</span></span>

|<span data-ttu-id="9b68d-167">Требование</span><span class="sxs-lookup"><span data-stu-id="9b68d-167">Requirement</span></span>| <span data-ttu-id="9b68d-168">Значение</span><span class="sxs-lookup"><span data-stu-id="9b68d-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b68d-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9b68d-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b68d-170">1.0</span><span class="sxs-lookup"><span data-stu-id="9b68d-170">1.0</span></span>|
|[<span data-ttu-id="9b68d-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9b68d-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b68d-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9b68d-172">Compose or Read</span></span>|
