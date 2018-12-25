---
title: Пространство имен Office — набор обязательных элементов 1.1
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 1670e8a1f2956979d31a7ce172bbfe81280addbb
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433847"
---
# <a name="office"></a><span data-ttu-id="c25e9-102">Office</span><span class="sxs-lookup"><span data-stu-id="c25e9-102">Office</span></span>

<span data-ttu-id="c25e9-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c25e9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c25e9-105">Требования</span><span class="sxs-lookup"><span data-stu-id="c25e9-105">Requirements</span></span>

|<span data-ttu-id="c25e9-106">Требование</span><span class="sxs-lookup"><span data-stu-id="c25e9-106">Requirement</span></span>| <span data-ttu-id="c25e9-107">Значение</span><span class="sxs-lookup"><span data-stu-id="c25e9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c25e9-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c25e9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c25e9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c25e9-109">1.0</span></span>|
|[<span data-ttu-id="c25e9-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c25e9-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c25e9-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c25e9-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="c25e9-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="c25e9-112">Namespaces</span></span>

<span data-ttu-id="c25e9-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="c25e9-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c25e9-114">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c25e9-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c25e9-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="c25e9-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c25e9-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c25e9-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="c25e9-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="c25e9-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c25e9-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="c25e9-118">Type:</span></span>

*   <span data-ttu-id="c25e9-119">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c25e9-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c25e9-120">Properties:</span></span>

|<span data-ttu-id="c25e9-121">Имя</span><span class="sxs-lookup"><span data-stu-id="c25e9-121">Name</span></span>| <span data-ttu-id="c25e9-122">Тип</span><span class="sxs-lookup"><span data-stu-id="c25e9-122">Type</span></span>| <span data-ttu-id="c25e9-123">Описание</span><span class="sxs-lookup"><span data-stu-id="c25e9-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c25e9-124">Для указания</span><span class="sxs-lookup"><span data-stu-id="c25e9-124">String</span></span>|<span data-ttu-id="c25e9-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="c25e9-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c25e9-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="c25e9-126">String</span></span>|<span data-ttu-id="c25e9-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="c25e9-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c25e9-128">Требования</span><span class="sxs-lookup"><span data-stu-id="c25e9-128">Requirements</span></span>

|<span data-ttu-id="c25e9-129">Требование</span><span class="sxs-lookup"><span data-stu-id="c25e9-129">Requirement</span></span>| <span data-ttu-id="c25e9-130">Значение</span><span class="sxs-lookup"><span data-stu-id="c25e9-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="c25e9-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c25e9-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c25e9-132">1.0</span><span class="sxs-lookup"><span data-stu-id="c25e9-132">1.0</span></span>|
|[<span data-ttu-id="c25e9-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c25e9-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c25e9-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c25e9-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="c25e9-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c25e9-135">CoercionType :String</span></span>

<span data-ttu-id="c25e9-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c25e9-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c25e9-137">Тип:</span><span class="sxs-lookup"><span data-stu-id="c25e9-137">Type:</span></span>

*   <span data-ttu-id="c25e9-138">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c25e9-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c25e9-139">Properties:</span></span>

|<span data-ttu-id="c25e9-140">Имя</span><span class="sxs-lookup"><span data-stu-id="c25e9-140">Name</span></span>| <span data-ttu-id="c25e9-141">Тип</span><span class="sxs-lookup"><span data-stu-id="c25e9-141">Type</span></span>| <span data-ttu-id="c25e9-142">Описание</span><span class="sxs-lookup"><span data-stu-id="c25e9-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c25e9-143">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-143">String</span></span>|<span data-ttu-id="c25e9-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c25e9-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c25e9-145">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-145">String</span></span>|<span data-ttu-id="c25e9-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="c25e9-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c25e9-147">Требования</span><span class="sxs-lookup"><span data-stu-id="c25e9-147">Requirements</span></span>

|<span data-ttu-id="c25e9-148">Требование</span><span class="sxs-lookup"><span data-stu-id="c25e9-148">Requirement</span></span>| <span data-ttu-id="c25e9-149">Значение</span><span class="sxs-lookup"><span data-stu-id="c25e9-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="c25e9-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c25e9-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c25e9-151">1.0</span><span class="sxs-lookup"><span data-stu-id="c25e9-151">1.0</span></span>|
|[<span data-ttu-id="c25e9-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c25e9-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c25e9-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c25e9-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="c25e9-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c25e9-154">SourceProperty :String</span></span>

<span data-ttu-id="c25e9-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c25e9-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c25e9-156">Тип:</span><span class="sxs-lookup"><span data-stu-id="c25e9-156">Type:</span></span>

*   <span data-ttu-id="c25e9-157">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c25e9-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c25e9-158">Properties:</span></span>

|<span data-ttu-id="c25e9-159">Имя</span><span class="sxs-lookup"><span data-stu-id="c25e9-159">Name</span></span>| <span data-ttu-id="c25e9-160">Тип</span><span class="sxs-lookup"><span data-stu-id="c25e9-160">Type</span></span>| <span data-ttu-id="c25e9-161">Описание</span><span class="sxs-lookup"><span data-stu-id="c25e9-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c25e9-162">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-162">String</span></span>|<span data-ttu-id="c25e9-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="c25e9-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c25e9-164">String</span><span class="sxs-lookup"><span data-stu-id="c25e9-164">String</span></span>|<span data-ttu-id="c25e9-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="c25e9-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c25e9-166">Требования</span><span class="sxs-lookup"><span data-stu-id="c25e9-166">Requirements</span></span>

|<span data-ttu-id="c25e9-167">Требование</span><span class="sxs-lookup"><span data-stu-id="c25e9-167">Requirement</span></span>| <span data-ttu-id="c25e9-168">Значение</span><span class="sxs-lookup"><span data-stu-id="c25e9-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="c25e9-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c25e9-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c25e9-170">1.0</span><span class="sxs-lookup"><span data-stu-id="c25e9-170">1.0</span></span>|
|[<span data-ttu-id="c25e9-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c25e9-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c25e9-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c25e9-172">Compose or read</span></span>|