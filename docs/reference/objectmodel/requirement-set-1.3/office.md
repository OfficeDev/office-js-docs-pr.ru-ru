---
title: Пространство имен Office — набор обязательных элементов 1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ef01b7da3d447af852a5558853e0902eab815dd3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451901"
---
# <a name="office"></a><span data-ttu-id="373a0-102">Office</span><span class="sxs-lookup"><span data-stu-id="373a0-102">Office</span></span>

<span data-ttu-id="373a0-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="373a0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="373a0-105">Требования</span><span class="sxs-lookup"><span data-stu-id="373a0-105">Requirements</span></span>

|<span data-ttu-id="373a0-106">Требование</span><span class="sxs-lookup"><span data-stu-id="373a0-106">Requirement</span></span>| <span data-ttu-id="373a0-107">Значение</span><span class="sxs-lookup"><span data-stu-id="373a0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="373a0-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="373a0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="373a0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="373a0-109">1.0</span></span>|
|[<span data-ttu-id="373a0-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="373a0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="373a0-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="373a0-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="373a0-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="373a0-112">Namespaces</span></span>

<span data-ttu-id="373a0-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="373a0-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="373a0-114">[MailboxEnums.](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="373a0-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="373a0-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="373a0-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="373a0-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="373a0-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="373a0-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="373a0-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="373a0-118">Тип</span><span class="sxs-lookup"><span data-stu-id="373a0-118">Type</span></span>

*   <span data-ttu-id="373a0-119">String</span><span class="sxs-lookup"><span data-stu-id="373a0-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="373a0-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="373a0-120">Properties:</span></span>

|<span data-ttu-id="373a0-121">Имя</span><span class="sxs-lookup"><span data-stu-id="373a0-121">Name</span></span>| <span data-ttu-id="373a0-122">Тип</span><span class="sxs-lookup"><span data-stu-id="373a0-122">Type</span></span>| <span data-ttu-id="373a0-123">Описание</span><span class="sxs-lookup"><span data-stu-id="373a0-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="373a0-124">Строка</span><span class="sxs-lookup"><span data-stu-id="373a0-124">String</span></span>|<span data-ttu-id="373a0-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="373a0-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="373a0-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="373a0-126">String</span></span>|<span data-ttu-id="373a0-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="373a0-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="373a0-128">Требования</span><span class="sxs-lookup"><span data-stu-id="373a0-128">Requirements</span></span>

|<span data-ttu-id="373a0-129">Требование</span><span class="sxs-lookup"><span data-stu-id="373a0-129">Requirement</span></span>| <span data-ttu-id="373a0-130">Значение</span><span class="sxs-lookup"><span data-stu-id="373a0-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="373a0-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="373a0-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="373a0-132">1.0</span><span class="sxs-lookup"><span data-stu-id="373a0-132">1.0</span></span>|
|[<span data-ttu-id="373a0-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="373a0-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="373a0-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="373a0-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="373a0-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="373a0-135">CoercionType :String</span></span>

<span data-ttu-id="373a0-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="373a0-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="373a0-137">Тип</span><span class="sxs-lookup"><span data-stu-id="373a0-137">Type</span></span>

*   <span data-ttu-id="373a0-138">String</span><span class="sxs-lookup"><span data-stu-id="373a0-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="373a0-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="373a0-139">Properties:</span></span>

|<span data-ttu-id="373a0-140">Имя</span><span class="sxs-lookup"><span data-stu-id="373a0-140">Name</span></span>| <span data-ttu-id="373a0-141">Тип</span><span class="sxs-lookup"><span data-stu-id="373a0-141">Type</span></span>| <span data-ttu-id="373a0-142">Описание</span><span class="sxs-lookup"><span data-stu-id="373a0-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="373a0-143">Строка</span><span class="sxs-lookup"><span data-stu-id="373a0-143">String</span></span>|<span data-ttu-id="373a0-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="373a0-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="373a0-145">Строка</span><span class="sxs-lookup"><span data-stu-id="373a0-145">String</span></span>|<span data-ttu-id="373a0-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="373a0-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="373a0-147">Требования</span><span class="sxs-lookup"><span data-stu-id="373a0-147">Requirements</span></span>

|<span data-ttu-id="373a0-148">Требование</span><span class="sxs-lookup"><span data-stu-id="373a0-148">Requirement</span></span>| <span data-ttu-id="373a0-149">Значение</span><span class="sxs-lookup"><span data-stu-id="373a0-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="373a0-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="373a0-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="373a0-151">1.0</span><span class="sxs-lookup"><span data-stu-id="373a0-151">1.0</span></span>|
|[<span data-ttu-id="373a0-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="373a0-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="373a0-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="373a0-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="373a0-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="373a0-154">SourceProperty :String</span></span>

<span data-ttu-id="373a0-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="373a0-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="373a0-156">Тип</span><span class="sxs-lookup"><span data-stu-id="373a0-156">Type</span></span>

*   <span data-ttu-id="373a0-157">String</span><span class="sxs-lookup"><span data-stu-id="373a0-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="373a0-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="373a0-158">Properties:</span></span>

|<span data-ttu-id="373a0-159">Имя</span><span class="sxs-lookup"><span data-stu-id="373a0-159">Name</span></span>| <span data-ttu-id="373a0-160">Тип</span><span class="sxs-lookup"><span data-stu-id="373a0-160">Type</span></span>| <span data-ttu-id="373a0-161">Описание</span><span class="sxs-lookup"><span data-stu-id="373a0-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="373a0-162">Строка</span><span class="sxs-lookup"><span data-stu-id="373a0-162">String</span></span>|<span data-ttu-id="373a0-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="373a0-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="373a0-164">Строка</span><span class="sxs-lookup"><span data-stu-id="373a0-164">String</span></span>|<span data-ttu-id="373a0-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="373a0-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="373a0-166">Требования</span><span class="sxs-lookup"><span data-stu-id="373a0-166">Requirements</span></span>

|<span data-ttu-id="373a0-167">Требование</span><span class="sxs-lookup"><span data-stu-id="373a0-167">Requirement</span></span>| <span data-ttu-id="373a0-168">Значение</span><span class="sxs-lookup"><span data-stu-id="373a0-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="373a0-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="373a0-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="373a0-170">1.0</span><span class="sxs-lookup"><span data-stu-id="373a0-170">1.0</span></span>|
|[<span data-ttu-id="373a0-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="373a0-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="373a0-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="373a0-172">Compose or Read</span></span>|
