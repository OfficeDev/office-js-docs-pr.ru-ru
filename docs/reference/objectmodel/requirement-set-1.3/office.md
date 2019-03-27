---
title: Пространство имен Office — набор обязательных элементов 1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ef01b7da3d447af852a5558853e0902eab815dd3
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871223"
---
# <a name="office"></a><span data-ttu-id="f7ff3-102">Office</span><span class="sxs-lookup"><span data-stu-id="f7ff3-102">Office</span></span>

<span data-ttu-id="f7ff3-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f7ff3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7ff3-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f7ff3-105">Requirements</span></span>

|<span data-ttu-id="f7ff3-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f7ff3-106">Requirement</span></span>| <span data-ttu-id="f7ff3-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7ff3-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7ff3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7ff3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f7ff3-109">1.0</span></span>|
|[<span data-ttu-id="f7ff3-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7ff3-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7ff3-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="f7ff3-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f7ff3-112">Namespaces</span></span>

<span data-ttu-id="f7ff3-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f7ff3-114">[MailboxEnums.](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f7ff3-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="f7ff3-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f7ff3-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="f7ff3-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f7ff3-118">Тип</span><span class="sxs-lookup"><span data-stu-id="f7ff3-118">Type</span></span>

*   <span data-ttu-id="f7ff3-119">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7ff3-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7ff3-120">Properties:</span></span>

|<span data-ttu-id="f7ff3-121">Имя</span><span class="sxs-lookup"><span data-stu-id="f7ff3-121">Name</span></span>| <span data-ttu-id="f7ff3-122">Тип</span><span class="sxs-lookup"><span data-stu-id="f7ff3-122">Type</span></span>| <span data-ttu-id="f7ff3-123">Описание</span><span class="sxs-lookup"><span data-stu-id="f7ff3-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f7ff3-124">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-124">String</span></span>|<span data-ttu-id="f7ff3-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f7ff3-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="f7ff3-126">String</span></span>|<span data-ttu-id="f7ff3-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7ff3-128">Требования</span><span class="sxs-lookup"><span data-stu-id="f7ff3-128">Requirements</span></span>

|<span data-ttu-id="f7ff3-129">Требование</span><span class="sxs-lookup"><span data-stu-id="f7ff3-129">Requirement</span></span>| <span data-ttu-id="f7ff3-130">Значение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7ff3-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7ff3-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7ff3-132">1.0</span><span class="sxs-lookup"><span data-stu-id="f7ff3-132">1.0</span></span>|
|[<span data-ttu-id="f7ff3-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7ff3-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7ff3-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="f7ff3-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-135">CoercionType :String</span></span>

<span data-ttu-id="f7ff3-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7ff3-137">Тип</span><span class="sxs-lookup"><span data-stu-id="f7ff3-137">Type</span></span>

*   <span data-ttu-id="f7ff3-138">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7ff3-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7ff3-139">Properties:</span></span>

|<span data-ttu-id="f7ff3-140">Имя</span><span class="sxs-lookup"><span data-stu-id="f7ff3-140">Name</span></span>| <span data-ttu-id="f7ff3-141">Тип</span><span class="sxs-lookup"><span data-stu-id="f7ff3-141">Type</span></span>| <span data-ttu-id="f7ff3-142">Описание</span><span class="sxs-lookup"><span data-stu-id="f7ff3-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f7ff3-143">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-143">String</span></span>|<span data-ttu-id="f7ff3-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f7ff3-145">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-145">String</span></span>|<span data-ttu-id="f7ff3-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7ff3-147">Требования</span><span class="sxs-lookup"><span data-stu-id="f7ff3-147">Requirements</span></span>

|<span data-ttu-id="f7ff3-148">Требование</span><span class="sxs-lookup"><span data-stu-id="f7ff3-148">Requirement</span></span>| <span data-ttu-id="f7ff3-149">Значение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7ff3-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7ff3-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7ff3-151">1.0</span><span class="sxs-lookup"><span data-stu-id="f7ff3-151">1.0</span></span>|
|[<span data-ttu-id="f7ff3-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7ff3-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7ff3-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="f7ff3-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-154">SourceProperty :String</span></span>

<span data-ttu-id="f7ff3-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7ff3-156">Тип</span><span class="sxs-lookup"><span data-stu-id="f7ff3-156">Type</span></span>

*   <span data-ttu-id="f7ff3-157">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7ff3-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7ff3-158">Properties:</span></span>

|<span data-ttu-id="f7ff3-159">Имя</span><span class="sxs-lookup"><span data-stu-id="f7ff3-159">Name</span></span>| <span data-ttu-id="f7ff3-160">Тип</span><span class="sxs-lookup"><span data-stu-id="f7ff3-160">Type</span></span>| <span data-ttu-id="f7ff3-161">Описание</span><span class="sxs-lookup"><span data-stu-id="f7ff3-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f7ff3-162">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-162">String</span></span>|<span data-ttu-id="f7ff3-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f7ff3-164">String</span><span class="sxs-lookup"><span data-stu-id="f7ff3-164">String</span></span>|<span data-ttu-id="f7ff3-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="f7ff3-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7ff3-166">Требования</span><span class="sxs-lookup"><span data-stu-id="f7ff3-166">Requirements</span></span>

|<span data-ttu-id="f7ff3-167">Требование</span><span class="sxs-lookup"><span data-stu-id="f7ff3-167">Requirement</span></span>| <span data-ttu-id="f7ff3-168">Значение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7ff3-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7ff3-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7ff3-170">1.0</span><span class="sxs-lookup"><span data-stu-id="f7ff3-170">1.0</span></span>|
|[<span data-ttu-id="f7ff3-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7ff3-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7ff3-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7ff3-172">Compose or Read</span></span>|
