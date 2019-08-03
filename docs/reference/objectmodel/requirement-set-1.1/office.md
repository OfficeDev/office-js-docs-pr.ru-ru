---
title: Пространство имен Office — набор обязательных элементов 1,1
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 01e60e0b01c7ca98bdff2c99dfcbe8944f822b81
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064363"
---
# <a name="office"></a><span data-ttu-id="637a7-102">Office</span><span class="sxs-lookup"><span data-stu-id="637a7-102">Office</span></span>

<span data-ttu-id="637a7-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="637a7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="637a7-105">Требования</span><span class="sxs-lookup"><span data-stu-id="637a7-105">Requirements</span></span>

|<span data-ttu-id="637a7-106">Требование</span><span class="sxs-lookup"><span data-stu-id="637a7-106">Requirement</span></span>| <span data-ttu-id="637a7-107">Значение</span><span class="sxs-lookup"><span data-stu-id="637a7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="637a7-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637a7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637a7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="637a7-109">1.0</span></span>|
|[<span data-ttu-id="637a7-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637a7-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637a7-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637a7-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="637a7-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="637a7-112">Namespaces</span></span>

<span data-ttu-id="637a7-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="637a7-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="637a7-114">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="637a7-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="637a7-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="637a7-115">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="637a7-116">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="637a7-116">AsyncResultStatus: String</span></span>

<span data-ttu-id="637a7-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="637a7-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="637a7-118">Тип</span><span class="sxs-lookup"><span data-stu-id="637a7-118">Type</span></span>

*   <span data-ttu-id="637a7-119">String</span><span class="sxs-lookup"><span data-stu-id="637a7-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="637a7-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="637a7-120">Properties:</span></span>

|<span data-ttu-id="637a7-121">Имя</span><span class="sxs-lookup"><span data-stu-id="637a7-121">Name</span></span>| <span data-ttu-id="637a7-122">Тип</span><span class="sxs-lookup"><span data-stu-id="637a7-122">Type</span></span>| <span data-ttu-id="637a7-123">Описание</span><span class="sxs-lookup"><span data-stu-id="637a7-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="637a7-124">String</span><span class="sxs-lookup"><span data-stu-id="637a7-124">String</span></span>|<span data-ttu-id="637a7-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="637a7-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="637a7-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="637a7-126">String</span></span>|<span data-ttu-id="637a7-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="637a7-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="637a7-128">Требования</span><span class="sxs-lookup"><span data-stu-id="637a7-128">Requirements</span></span>

|<span data-ttu-id="637a7-129">Требование</span><span class="sxs-lookup"><span data-stu-id="637a7-129">Requirement</span></span>| <span data-ttu-id="637a7-130">Значение</span><span class="sxs-lookup"><span data-stu-id="637a7-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="637a7-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637a7-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637a7-132">1.0</span><span class="sxs-lookup"><span data-stu-id="637a7-132">1.0</span></span>|
|[<span data-ttu-id="637a7-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637a7-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637a7-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637a7-134">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="637a7-135">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="637a7-135">CoercionType: String</span></span>

<span data-ttu-id="637a7-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="637a7-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="637a7-137">Тип</span><span class="sxs-lookup"><span data-stu-id="637a7-137">Type</span></span>

*   <span data-ttu-id="637a7-138">String</span><span class="sxs-lookup"><span data-stu-id="637a7-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="637a7-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="637a7-139">Properties:</span></span>

|<span data-ttu-id="637a7-140">Имя</span><span class="sxs-lookup"><span data-stu-id="637a7-140">Name</span></span>| <span data-ttu-id="637a7-141">Тип</span><span class="sxs-lookup"><span data-stu-id="637a7-141">Type</span></span>| <span data-ttu-id="637a7-142">Описание</span><span class="sxs-lookup"><span data-stu-id="637a7-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="637a7-143">String</span><span class="sxs-lookup"><span data-stu-id="637a7-143">String</span></span>|<span data-ttu-id="637a7-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="637a7-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="637a7-145">String</span><span class="sxs-lookup"><span data-stu-id="637a7-145">String</span></span>|<span data-ttu-id="637a7-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="637a7-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="637a7-147">Требования</span><span class="sxs-lookup"><span data-stu-id="637a7-147">Requirements</span></span>

|<span data-ttu-id="637a7-148">Требование</span><span class="sxs-lookup"><span data-stu-id="637a7-148">Requirement</span></span>| <span data-ttu-id="637a7-149">Значение</span><span class="sxs-lookup"><span data-stu-id="637a7-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="637a7-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637a7-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637a7-151">1.0</span><span class="sxs-lookup"><span data-stu-id="637a7-151">1.0</span></span>|
|[<span data-ttu-id="637a7-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637a7-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637a7-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637a7-153">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="637a7-154">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="637a7-154">SourceProperty: String</span></span>

<span data-ttu-id="637a7-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="637a7-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="637a7-156">Тип</span><span class="sxs-lookup"><span data-stu-id="637a7-156">Type</span></span>

*   <span data-ttu-id="637a7-157">String</span><span class="sxs-lookup"><span data-stu-id="637a7-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="637a7-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="637a7-158">Properties:</span></span>

|<span data-ttu-id="637a7-159">Имя</span><span class="sxs-lookup"><span data-stu-id="637a7-159">Name</span></span>| <span data-ttu-id="637a7-160">Тип</span><span class="sxs-lookup"><span data-stu-id="637a7-160">Type</span></span>| <span data-ttu-id="637a7-161">Описание</span><span class="sxs-lookup"><span data-stu-id="637a7-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="637a7-162">String</span><span class="sxs-lookup"><span data-stu-id="637a7-162">String</span></span>|<span data-ttu-id="637a7-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="637a7-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="637a7-164">String</span><span class="sxs-lookup"><span data-stu-id="637a7-164">String</span></span>|<span data-ttu-id="637a7-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="637a7-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="637a7-166">Требования</span><span class="sxs-lookup"><span data-stu-id="637a7-166">Requirements</span></span>

|<span data-ttu-id="637a7-167">Требование</span><span class="sxs-lookup"><span data-stu-id="637a7-167">Requirement</span></span>| <span data-ttu-id="637a7-168">Значение</span><span class="sxs-lookup"><span data-stu-id="637a7-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="637a7-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637a7-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637a7-170">1.0</span><span class="sxs-lookup"><span data-stu-id="637a7-170">1.0</span></span>|
|[<span data-ttu-id="637a7-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637a7-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637a7-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637a7-172">Compose or Read</span></span>|
