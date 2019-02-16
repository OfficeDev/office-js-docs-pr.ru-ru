---
title: Пространство имен Office — набор обязательных элементов 1.4
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: af5e05e2243c0132018bc4eba7006f9a5aad4099
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067953"
---
# <a name="office"></a><span data-ttu-id="c60bc-102">Office</span><span class="sxs-lookup"><span data-stu-id="c60bc-102">Office</span></span>

<span data-ttu-id="c60bc-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c60bc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c60bc-105">Требования</span><span class="sxs-lookup"><span data-stu-id="c60bc-105">Requirements</span></span>

|<span data-ttu-id="c60bc-106">Требование</span><span class="sxs-lookup"><span data-stu-id="c60bc-106">Requirement</span></span>| <span data-ttu-id="c60bc-107">Значение</span><span class="sxs-lookup"><span data-stu-id="c60bc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c60bc-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c60bc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c60bc-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c60bc-109">1.0</span></span>|
|[<span data-ttu-id="c60bc-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c60bc-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c60bc-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c60bc-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="c60bc-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="c60bc-112">Namespaces</span></span>

<span data-ttu-id="c60bc-113">[context.](Office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="c60bc-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c60bc-114">[MailboxEnums.](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c60bc-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c60bc-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="c60bc-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c60bc-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c60bc-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="c60bc-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="c60bc-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c60bc-118">Тип</span><span class="sxs-lookup"><span data-stu-id="c60bc-118">Type</span></span>

*   <span data-ttu-id="c60bc-119">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c60bc-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c60bc-120">Properties:</span></span>

|<span data-ttu-id="c60bc-121">Имя</span><span class="sxs-lookup"><span data-stu-id="c60bc-121">Name</span></span>| <span data-ttu-id="c60bc-122">Тип</span><span class="sxs-lookup"><span data-stu-id="c60bc-122">Type</span></span>| <span data-ttu-id="c60bc-123">Описание</span><span class="sxs-lookup"><span data-stu-id="c60bc-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c60bc-124">Для указания</span><span class="sxs-lookup"><span data-stu-id="c60bc-124">String</span></span>|<span data-ttu-id="c60bc-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="c60bc-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c60bc-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="c60bc-126">String</span></span>|<span data-ttu-id="c60bc-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="c60bc-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c60bc-128">Требования</span><span class="sxs-lookup"><span data-stu-id="c60bc-128">Requirements</span></span>

|<span data-ttu-id="c60bc-129">Требование</span><span class="sxs-lookup"><span data-stu-id="c60bc-129">Requirement</span></span>| <span data-ttu-id="c60bc-130">Значение</span><span class="sxs-lookup"><span data-stu-id="c60bc-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="c60bc-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c60bc-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c60bc-132">1.0</span><span class="sxs-lookup"><span data-stu-id="c60bc-132">1.0</span></span>|
|[<span data-ttu-id="c60bc-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c60bc-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c60bc-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c60bc-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="c60bc-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c60bc-135">CoercionType :String</span></span>

<span data-ttu-id="c60bc-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c60bc-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c60bc-137">Тип</span><span class="sxs-lookup"><span data-stu-id="c60bc-137">Type</span></span>

*   <span data-ttu-id="c60bc-138">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c60bc-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c60bc-139">Properties:</span></span>

|<span data-ttu-id="c60bc-140">Имя</span><span class="sxs-lookup"><span data-stu-id="c60bc-140">Name</span></span>| <span data-ttu-id="c60bc-141">Тип</span><span class="sxs-lookup"><span data-stu-id="c60bc-141">Type</span></span>| <span data-ttu-id="c60bc-142">Описание</span><span class="sxs-lookup"><span data-stu-id="c60bc-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c60bc-143">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-143">String</span></span>|<span data-ttu-id="c60bc-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c60bc-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c60bc-145">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-145">String</span></span>|<span data-ttu-id="c60bc-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="c60bc-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c60bc-147">Требования</span><span class="sxs-lookup"><span data-stu-id="c60bc-147">Requirements</span></span>

|<span data-ttu-id="c60bc-148">Требование</span><span class="sxs-lookup"><span data-stu-id="c60bc-148">Requirement</span></span>| <span data-ttu-id="c60bc-149">Значение</span><span class="sxs-lookup"><span data-stu-id="c60bc-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="c60bc-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c60bc-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c60bc-151">1.0</span><span class="sxs-lookup"><span data-stu-id="c60bc-151">1.0</span></span>|
|[<span data-ttu-id="c60bc-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c60bc-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c60bc-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c60bc-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="c60bc-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c60bc-154">SourceProperty :String</span></span>

<span data-ttu-id="c60bc-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c60bc-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c60bc-156">Тип</span><span class="sxs-lookup"><span data-stu-id="c60bc-156">Type</span></span>

*   <span data-ttu-id="c60bc-157">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c60bc-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c60bc-158">Properties:</span></span>

|<span data-ttu-id="c60bc-159">Имя</span><span class="sxs-lookup"><span data-stu-id="c60bc-159">Name</span></span>| <span data-ttu-id="c60bc-160">Тип</span><span class="sxs-lookup"><span data-stu-id="c60bc-160">Type</span></span>| <span data-ttu-id="c60bc-161">Описание</span><span class="sxs-lookup"><span data-stu-id="c60bc-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c60bc-162">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-162">String</span></span>|<span data-ttu-id="c60bc-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="c60bc-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c60bc-164">String</span><span class="sxs-lookup"><span data-stu-id="c60bc-164">String</span></span>|<span data-ttu-id="c60bc-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="c60bc-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c60bc-166">Требования</span><span class="sxs-lookup"><span data-stu-id="c60bc-166">Requirements</span></span>

|<span data-ttu-id="c60bc-167">Требование</span><span class="sxs-lookup"><span data-stu-id="c60bc-167">Requirement</span></span>| <span data-ttu-id="c60bc-168">Значение</span><span class="sxs-lookup"><span data-stu-id="c60bc-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="c60bc-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c60bc-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c60bc-170">1.0</span><span class="sxs-lookup"><span data-stu-id="c60bc-170">1.0</span></span>|
|[<span data-ttu-id="c60bc-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c60bc-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c60bc-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c60bc-172">Compose or Read</span></span>|
