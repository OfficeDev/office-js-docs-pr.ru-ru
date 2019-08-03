---
title: Пространство имен Office — набор обязательных элементов 1,2
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 9dd492046df6325c5c2cdb04dbd1c8bc331b3471
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064398"
---
# <a name="office"></a><span data-ttu-id="637ed-102">Office</span><span class="sxs-lookup"><span data-stu-id="637ed-102">Office</span></span>

<span data-ttu-id="637ed-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="637ed-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="637ed-105">Требования</span><span class="sxs-lookup"><span data-stu-id="637ed-105">Requirements</span></span>

|<span data-ttu-id="637ed-106">Требование</span><span class="sxs-lookup"><span data-stu-id="637ed-106">Requirement</span></span>| <span data-ttu-id="637ed-107">Значение</span><span class="sxs-lookup"><span data-stu-id="637ed-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="637ed-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637ed-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637ed-109">1.0</span><span class="sxs-lookup"><span data-stu-id="637ed-109">1.0</span></span>|
|[<span data-ttu-id="637ed-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637ed-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637ed-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637ed-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="637ed-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="637ed-112">Namespaces</span></span>

<span data-ttu-id="637ed-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="637ed-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="637ed-114">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="637ed-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="637ed-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="637ed-115">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="637ed-116">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="637ed-116">AsyncResultStatus: String</span></span>

<span data-ttu-id="637ed-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="637ed-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="637ed-118">Тип</span><span class="sxs-lookup"><span data-stu-id="637ed-118">Type</span></span>

*   <span data-ttu-id="637ed-119">String</span><span class="sxs-lookup"><span data-stu-id="637ed-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="637ed-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="637ed-120">Properties:</span></span>

|<span data-ttu-id="637ed-121">Имя</span><span class="sxs-lookup"><span data-stu-id="637ed-121">Name</span></span>| <span data-ttu-id="637ed-122">Тип</span><span class="sxs-lookup"><span data-stu-id="637ed-122">Type</span></span>| <span data-ttu-id="637ed-123">Описание</span><span class="sxs-lookup"><span data-stu-id="637ed-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="637ed-124">String</span><span class="sxs-lookup"><span data-stu-id="637ed-124">String</span></span>|<span data-ttu-id="637ed-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="637ed-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="637ed-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="637ed-126">String</span></span>|<span data-ttu-id="637ed-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="637ed-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="637ed-128">Требования</span><span class="sxs-lookup"><span data-stu-id="637ed-128">Requirements</span></span>

|<span data-ttu-id="637ed-129">Требование</span><span class="sxs-lookup"><span data-stu-id="637ed-129">Requirement</span></span>| <span data-ttu-id="637ed-130">Значение</span><span class="sxs-lookup"><span data-stu-id="637ed-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="637ed-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637ed-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637ed-132">1.0</span><span class="sxs-lookup"><span data-stu-id="637ed-132">1.0</span></span>|
|[<span data-ttu-id="637ed-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637ed-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637ed-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637ed-134">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="637ed-135">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="637ed-135">CoercionType: String</span></span>

<span data-ttu-id="637ed-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="637ed-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="637ed-137">Тип</span><span class="sxs-lookup"><span data-stu-id="637ed-137">Type</span></span>

*   <span data-ttu-id="637ed-138">String</span><span class="sxs-lookup"><span data-stu-id="637ed-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="637ed-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="637ed-139">Properties:</span></span>

|<span data-ttu-id="637ed-140">Имя</span><span class="sxs-lookup"><span data-stu-id="637ed-140">Name</span></span>| <span data-ttu-id="637ed-141">Тип</span><span class="sxs-lookup"><span data-stu-id="637ed-141">Type</span></span>| <span data-ttu-id="637ed-142">Описание</span><span class="sxs-lookup"><span data-stu-id="637ed-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="637ed-143">String</span><span class="sxs-lookup"><span data-stu-id="637ed-143">String</span></span>|<span data-ttu-id="637ed-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="637ed-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="637ed-145">String</span><span class="sxs-lookup"><span data-stu-id="637ed-145">String</span></span>|<span data-ttu-id="637ed-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="637ed-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="637ed-147">Требования</span><span class="sxs-lookup"><span data-stu-id="637ed-147">Requirements</span></span>

|<span data-ttu-id="637ed-148">Требование</span><span class="sxs-lookup"><span data-stu-id="637ed-148">Requirement</span></span>| <span data-ttu-id="637ed-149">Значение</span><span class="sxs-lookup"><span data-stu-id="637ed-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="637ed-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637ed-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637ed-151">1.0</span><span class="sxs-lookup"><span data-stu-id="637ed-151">1.0</span></span>|
|[<span data-ttu-id="637ed-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637ed-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637ed-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637ed-153">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="637ed-154">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="637ed-154">SourceProperty: String</span></span>

<span data-ttu-id="637ed-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="637ed-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="637ed-156">Тип</span><span class="sxs-lookup"><span data-stu-id="637ed-156">Type</span></span>

*   <span data-ttu-id="637ed-157">String</span><span class="sxs-lookup"><span data-stu-id="637ed-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="637ed-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="637ed-158">Properties:</span></span>

|<span data-ttu-id="637ed-159">Имя</span><span class="sxs-lookup"><span data-stu-id="637ed-159">Name</span></span>| <span data-ttu-id="637ed-160">Тип</span><span class="sxs-lookup"><span data-stu-id="637ed-160">Type</span></span>| <span data-ttu-id="637ed-161">Описание</span><span class="sxs-lookup"><span data-stu-id="637ed-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="637ed-162">String</span><span class="sxs-lookup"><span data-stu-id="637ed-162">String</span></span>|<span data-ttu-id="637ed-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="637ed-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="637ed-164">String</span><span class="sxs-lookup"><span data-stu-id="637ed-164">String</span></span>|<span data-ttu-id="637ed-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="637ed-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="637ed-166">Требования</span><span class="sxs-lookup"><span data-stu-id="637ed-166">Requirements</span></span>

|<span data-ttu-id="637ed-167">Требование</span><span class="sxs-lookup"><span data-stu-id="637ed-167">Requirement</span></span>| <span data-ttu-id="637ed-168">Значение</span><span class="sxs-lookup"><span data-stu-id="637ed-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="637ed-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="637ed-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="637ed-170">1.0</span><span class="sxs-lookup"><span data-stu-id="637ed-170">1.0</span></span>|
|[<span data-ttu-id="637ed-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="637ed-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="637ed-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="637ed-172">Compose or Read</span></span>|
