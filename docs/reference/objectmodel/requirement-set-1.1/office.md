---
title: Пространство имен Office — набор обязательных элементов 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: eda5e1fb5f2c11ae91e4a1479892c36ec23e1897
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871999"
---
# <a name="office"></a><span data-ttu-id="ffe8b-102">Office</span><span class="sxs-lookup"><span data-stu-id="ffe8b-102">Office</span></span>

<span data-ttu-id="ffe8b-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ffe8b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ffe8b-105">Требования</span><span class="sxs-lookup"><span data-stu-id="ffe8b-105">Requirements</span></span>

|<span data-ttu-id="ffe8b-106">Требование</span><span class="sxs-lookup"><span data-stu-id="ffe8b-106">Requirement</span></span>| <span data-ttu-id="ffe8b-107">Значение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffe8b-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ffe8b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ffe8b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ffe8b-109">1.0</span></span>|
|[<span data-ttu-id="ffe8b-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ffe8b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffe8b-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="ffe8b-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="ffe8b-112">Namespaces</span></span>

<span data-ttu-id="ffe8b-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ffe8b-114">[MailboxEnums.](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ffe8b-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="ffe8b-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ffe8b-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="ffe8b-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ffe8b-118">Тип</span><span class="sxs-lookup"><span data-stu-id="ffe8b-118">Type</span></span>

*   <span data-ttu-id="ffe8b-119">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ffe8b-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ffe8b-120">Properties:</span></span>

|<span data-ttu-id="ffe8b-121">Имя</span><span class="sxs-lookup"><span data-stu-id="ffe8b-121">Name</span></span>| <span data-ttu-id="ffe8b-122">Тип</span><span class="sxs-lookup"><span data-stu-id="ffe8b-122">Type</span></span>| <span data-ttu-id="ffe8b-123">Описание</span><span class="sxs-lookup"><span data-stu-id="ffe8b-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ffe8b-124">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-124">String</span></span>|<span data-ttu-id="ffe8b-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ffe8b-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="ffe8b-126">String</span></span>|<span data-ttu-id="ffe8b-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ffe8b-128">Требования</span><span class="sxs-lookup"><span data-stu-id="ffe8b-128">Requirements</span></span>

|<span data-ttu-id="ffe8b-129">Требование</span><span class="sxs-lookup"><span data-stu-id="ffe8b-129">Requirement</span></span>| <span data-ttu-id="ffe8b-130">Значение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffe8b-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ffe8b-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ffe8b-132">1.0</span><span class="sxs-lookup"><span data-stu-id="ffe8b-132">1.0</span></span>|
|[<span data-ttu-id="ffe8b-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ffe8b-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffe8b-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="ffe8b-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-135">CoercionType :String</span></span>

<span data-ttu-id="ffe8b-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ffe8b-137">Тип</span><span class="sxs-lookup"><span data-stu-id="ffe8b-137">Type</span></span>

*   <span data-ttu-id="ffe8b-138">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ffe8b-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ffe8b-139">Properties:</span></span>

|<span data-ttu-id="ffe8b-140">Имя</span><span class="sxs-lookup"><span data-stu-id="ffe8b-140">Name</span></span>| <span data-ttu-id="ffe8b-141">Тип</span><span class="sxs-lookup"><span data-stu-id="ffe8b-141">Type</span></span>| <span data-ttu-id="ffe8b-142">Описание</span><span class="sxs-lookup"><span data-stu-id="ffe8b-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ffe8b-143">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-143">String</span></span>|<span data-ttu-id="ffe8b-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ffe8b-145">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-145">String</span></span>|<span data-ttu-id="ffe8b-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ffe8b-147">Требования</span><span class="sxs-lookup"><span data-stu-id="ffe8b-147">Requirements</span></span>

|<span data-ttu-id="ffe8b-148">Требование</span><span class="sxs-lookup"><span data-stu-id="ffe8b-148">Requirement</span></span>| <span data-ttu-id="ffe8b-149">Значение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffe8b-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ffe8b-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ffe8b-151">1.0</span><span class="sxs-lookup"><span data-stu-id="ffe8b-151">1.0</span></span>|
|[<span data-ttu-id="ffe8b-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ffe8b-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffe8b-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="ffe8b-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-154">SourceProperty :String</span></span>

<span data-ttu-id="ffe8b-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ffe8b-156">Тип</span><span class="sxs-lookup"><span data-stu-id="ffe8b-156">Type</span></span>

*   <span data-ttu-id="ffe8b-157">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ffe8b-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ffe8b-158">Properties:</span></span>

|<span data-ttu-id="ffe8b-159">Имя</span><span class="sxs-lookup"><span data-stu-id="ffe8b-159">Name</span></span>| <span data-ttu-id="ffe8b-160">Тип</span><span class="sxs-lookup"><span data-stu-id="ffe8b-160">Type</span></span>| <span data-ttu-id="ffe8b-161">Описание</span><span class="sxs-lookup"><span data-stu-id="ffe8b-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ffe8b-162">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-162">String</span></span>|<span data-ttu-id="ffe8b-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ffe8b-164">String</span><span class="sxs-lookup"><span data-stu-id="ffe8b-164">String</span></span>|<span data-ttu-id="ffe8b-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="ffe8b-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ffe8b-166">Требования</span><span class="sxs-lookup"><span data-stu-id="ffe8b-166">Requirements</span></span>

|<span data-ttu-id="ffe8b-167">Требование</span><span class="sxs-lookup"><span data-stu-id="ffe8b-167">Requirement</span></span>| <span data-ttu-id="ffe8b-168">Значение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffe8b-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ffe8b-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ffe8b-170">1.0</span><span class="sxs-lookup"><span data-stu-id="ffe8b-170">1.0</span></span>|
|[<span data-ttu-id="ffe8b-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ffe8b-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffe8b-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ffe8b-172">Compose or Read</span></span>|
