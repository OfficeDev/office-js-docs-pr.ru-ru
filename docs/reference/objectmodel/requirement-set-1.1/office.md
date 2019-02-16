---
title: Пространство имен Office — набор обязательных элементов 1.1
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: b98f8fa01e2cfdf17d6105beab67199a2cec4317
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068023"
---
# <a name="office"></a><span data-ttu-id="eba9f-102">Office</span><span class="sxs-lookup"><span data-stu-id="eba9f-102">Office</span></span>

<span data-ttu-id="eba9f-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="eba9f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eba9f-105">Требования</span><span class="sxs-lookup"><span data-stu-id="eba9f-105">Requirements</span></span>

|<span data-ttu-id="eba9f-106">Требование</span><span class="sxs-lookup"><span data-stu-id="eba9f-106">Requirement</span></span>| <span data-ttu-id="eba9f-107">Значение</span><span class="sxs-lookup"><span data-stu-id="eba9f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba9f-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eba9f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba9f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="eba9f-109">1.0</span></span>|
|[<span data-ttu-id="eba9f-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eba9f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba9f-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eba9f-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="eba9f-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="eba9f-112">Namespaces</span></span>

<span data-ttu-id="eba9f-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="eba9f-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="eba9f-114">[MailboxEnums.](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="eba9f-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="eba9f-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="eba9f-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="eba9f-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="eba9f-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="eba9f-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="eba9f-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="eba9f-118">Тип</span><span class="sxs-lookup"><span data-stu-id="eba9f-118">Type</span></span>

*   <span data-ttu-id="eba9f-119">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eba9f-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="eba9f-120">Properties:</span></span>

|<span data-ttu-id="eba9f-121">Имя</span><span class="sxs-lookup"><span data-stu-id="eba9f-121">Name</span></span>| <span data-ttu-id="eba9f-122">Тип</span><span class="sxs-lookup"><span data-stu-id="eba9f-122">Type</span></span>| <span data-ttu-id="eba9f-123">Описание</span><span class="sxs-lookup"><span data-stu-id="eba9f-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="eba9f-124">Для указания</span><span class="sxs-lookup"><span data-stu-id="eba9f-124">String</span></span>|<span data-ttu-id="eba9f-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="eba9f-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="eba9f-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="eba9f-126">String</span></span>|<span data-ttu-id="eba9f-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="eba9f-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eba9f-128">Требования</span><span class="sxs-lookup"><span data-stu-id="eba9f-128">Requirements</span></span>

|<span data-ttu-id="eba9f-129">Требование</span><span class="sxs-lookup"><span data-stu-id="eba9f-129">Requirement</span></span>| <span data-ttu-id="eba9f-130">Значение</span><span class="sxs-lookup"><span data-stu-id="eba9f-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba9f-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eba9f-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba9f-132">1.0</span><span class="sxs-lookup"><span data-stu-id="eba9f-132">1.0</span></span>|
|[<span data-ttu-id="eba9f-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eba9f-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba9f-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eba9f-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="eba9f-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="eba9f-135">CoercionType :String</span></span>

<span data-ttu-id="eba9f-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="eba9f-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eba9f-137">Тип</span><span class="sxs-lookup"><span data-stu-id="eba9f-137">Type</span></span>

*   <span data-ttu-id="eba9f-138">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eba9f-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="eba9f-139">Properties:</span></span>

|<span data-ttu-id="eba9f-140">Имя</span><span class="sxs-lookup"><span data-stu-id="eba9f-140">Name</span></span>| <span data-ttu-id="eba9f-141">Тип</span><span class="sxs-lookup"><span data-stu-id="eba9f-141">Type</span></span>| <span data-ttu-id="eba9f-142">Описание</span><span class="sxs-lookup"><span data-stu-id="eba9f-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="eba9f-143">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-143">String</span></span>|<span data-ttu-id="eba9f-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="eba9f-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="eba9f-145">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-145">String</span></span>|<span data-ttu-id="eba9f-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="eba9f-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eba9f-147">Требования</span><span class="sxs-lookup"><span data-stu-id="eba9f-147">Requirements</span></span>

|<span data-ttu-id="eba9f-148">Требование</span><span class="sxs-lookup"><span data-stu-id="eba9f-148">Requirement</span></span>| <span data-ttu-id="eba9f-149">Значение</span><span class="sxs-lookup"><span data-stu-id="eba9f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba9f-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eba9f-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba9f-151">1.0</span><span class="sxs-lookup"><span data-stu-id="eba9f-151">1.0</span></span>|
|[<span data-ttu-id="eba9f-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eba9f-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba9f-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eba9f-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="eba9f-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="eba9f-154">SourceProperty :String</span></span>

<span data-ttu-id="eba9f-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="eba9f-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eba9f-156">Тип</span><span class="sxs-lookup"><span data-stu-id="eba9f-156">Type</span></span>

*   <span data-ttu-id="eba9f-157">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eba9f-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="eba9f-158">Properties:</span></span>

|<span data-ttu-id="eba9f-159">Имя</span><span class="sxs-lookup"><span data-stu-id="eba9f-159">Name</span></span>| <span data-ttu-id="eba9f-160">Тип</span><span class="sxs-lookup"><span data-stu-id="eba9f-160">Type</span></span>| <span data-ttu-id="eba9f-161">Описание</span><span class="sxs-lookup"><span data-stu-id="eba9f-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="eba9f-162">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-162">String</span></span>|<span data-ttu-id="eba9f-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="eba9f-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="eba9f-164">String</span><span class="sxs-lookup"><span data-stu-id="eba9f-164">String</span></span>|<span data-ttu-id="eba9f-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="eba9f-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eba9f-166">Требования</span><span class="sxs-lookup"><span data-stu-id="eba9f-166">Requirements</span></span>

|<span data-ttu-id="eba9f-167">Требование</span><span class="sxs-lookup"><span data-stu-id="eba9f-167">Requirement</span></span>| <span data-ttu-id="eba9f-168">Значение</span><span class="sxs-lookup"><span data-stu-id="eba9f-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba9f-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eba9f-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba9f-170">1.0</span><span class="sxs-lookup"><span data-stu-id="eba9f-170">1.0</span></span>|
|[<span data-ttu-id="eba9f-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eba9f-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba9f-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eba9f-172">Compose or Read</span></span>|
