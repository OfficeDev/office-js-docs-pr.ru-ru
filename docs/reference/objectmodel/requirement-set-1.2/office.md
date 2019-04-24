---
title: Пространство имен Office — набор обязательных элементов 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dc98d4c2da6e8f9ca294a6c686cf081478e1bb24
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450312"
---
# <a name="office"></a><span data-ttu-id="ad7b8-102">Office</span><span class="sxs-lookup"><span data-stu-id="ad7b8-102">Office</span></span>

<span data-ttu-id="ad7b8-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ad7b8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad7b8-105">Требования</span><span class="sxs-lookup"><span data-stu-id="ad7b8-105">Requirements</span></span>

|<span data-ttu-id="ad7b8-106">Требование</span><span class="sxs-lookup"><span data-stu-id="ad7b8-106">Requirement</span></span>| <span data-ttu-id="ad7b8-107">Значение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad7b8-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad7b8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad7b8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ad7b8-109">1.0</span></span>|
|[<span data-ttu-id="ad7b8-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad7b8-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad7b8-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="ad7b8-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="ad7b8-112">Namespaces</span></span>

<span data-ttu-id="ad7b8-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ad7b8-114">[MailboxEnums.](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ad7b8-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="ad7b8-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ad7b8-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ad7b8-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="ad7b8-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ad7b8-118">Тип</span><span class="sxs-lookup"><span data-stu-id="ad7b8-118">Type</span></span>

*   <span data-ttu-id="ad7b8-119">String</span><span class="sxs-lookup"><span data-stu-id="ad7b8-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad7b8-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad7b8-120">Properties:</span></span>

|<span data-ttu-id="ad7b8-121">Имя</span><span class="sxs-lookup"><span data-stu-id="ad7b8-121">Name</span></span>| <span data-ttu-id="ad7b8-122">Тип</span><span class="sxs-lookup"><span data-stu-id="ad7b8-122">Type</span></span>| <span data-ttu-id="ad7b8-123">Описание</span><span class="sxs-lookup"><span data-stu-id="ad7b8-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ad7b8-124">Строка</span><span class="sxs-lookup"><span data-stu-id="ad7b8-124">String</span></span>|<span data-ttu-id="ad7b8-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ad7b8-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="ad7b8-126">String</span></span>|<span data-ttu-id="ad7b8-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad7b8-128">Требования</span><span class="sxs-lookup"><span data-stu-id="ad7b8-128">Requirements</span></span>

|<span data-ttu-id="ad7b8-129">Требование</span><span class="sxs-lookup"><span data-stu-id="ad7b8-129">Requirement</span></span>| <span data-ttu-id="ad7b8-130">Значение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad7b8-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad7b8-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad7b8-132">1.0</span><span class="sxs-lookup"><span data-stu-id="ad7b8-132">1.0</span></span>|
|[<span data-ttu-id="ad7b8-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad7b8-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad7b8-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="ad7b8-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ad7b8-135">CoercionType :String</span></span>

<span data-ttu-id="ad7b8-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ad7b8-137">Тип</span><span class="sxs-lookup"><span data-stu-id="ad7b8-137">Type</span></span>

*   <span data-ttu-id="ad7b8-138">String</span><span class="sxs-lookup"><span data-stu-id="ad7b8-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad7b8-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad7b8-139">Properties:</span></span>

|<span data-ttu-id="ad7b8-140">Имя</span><span class="sxs-lookup"><span data-stu-id="ad7b8-140">Name</span></span>| <span data-ttu-id="ad7b8-141">Тип</span><span class="sxs-lookup"><span data-stu-id="ad7b8-141">Type</span></span>| <span data-ttu-id="ad7b8-142">Описание</span><span class="sxs-lookup"><span data-stu-id="ad7b8-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ad7b8-143">Строка</span><span class="sxs-lookup"><span data-stu-id="ad7b8-143">String</span></span>|<span data-ttu-id="ad7b8-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ad7b8-145">Строка</span><span class="sxs-lookup"><span data-stu-id="ad7b8-145">String</span></span>|<span data-ttu-id="ad7b8-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad7b8-147">Требования</span><span class="sxs-lookup"><span data-stu-id="ad7b8-147">Requirements</span></span>

|<span data-ttu-id="ad7b8-148">Требование</span><span class="sxs-lookup"><span data-stu-id="ad7b8-148">Requirement</span></span>| <span data-ttu-id="ad7b8-149">Значение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad7b8-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad7b8-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad7b8-151">1.0</span><span class="sxs-lookup"><span data-stu-id="ad7b8-151">1.0</span></span>|
|[<span data-ttu-id="ad7b8-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad7b8-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad7b8-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="ad7b8-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ad7b8-154">SourceProperty :String</span></span>

<span data-ttu-id="ad7b8-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ad7b8-156">Тип</span><span class="sxs-lookup"><span data-stu-id="ad7b8-156">Type</span></span>

*   <span data-ttu-id="ad7b8-157">String</span><span class="sxs-lookup"><span data-stu-id="ad7b8-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad7b8-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad7b8-158">Properties:</span></span>

|<span data-ttu-id="ad7b8-159">Имя</span><span class="sxs-lookup"><span data-stu-id="ad7b8-159">Name</span></span>| <span data-ttu-id="ad7b8-160">Тип</span><span class="sxs-lookup"><span data-stu-id="ad7b8-160">Type</span></span>| <span data-ttu-id="ad7b8-161">Описание</span><span class="sxs-lookup"><span data-stu-id="ad7b8-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ad7b8-162">Строка</span><span class="sxs-lookup"><span data-stu-id="ad7b8-162">String</span></span>|<span data-ttu-id="ad7b8-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ad7b8-164">Строка</span><span class="sxs-lookup"><span data-stu-id="ad7b8-164">String</span></span>|<span data-ttu-id="ad7b8-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad7b8-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad7b8-166">Требования</span><span class="sxs-lookup"><span data-stu-id="ad7b8-166">Requirements</span></span>

|<span data-ttu-id="ad7b8-167">Требование</span><span class="sxs-lookup"><span data-stu-id="ad7b8-167">Requirement</span></span>| <span data-ttu-id="ad7b8-168">Значение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad7b8-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad7b8-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad7b8-170">1.0</span><span class="sxs-lookup"><span data-stu-id="ad7b8-170">1.0</span></span>|
|[<span data-ttu-id="ad7b8-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad7b8-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad7b8-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad7b8-172">Compose or Read</span></span>|
