---
title: Пространство имен Office — набор обязательных элементов 1.2
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 0dfd20d8c45f7630f388cb7ecde9ffb2a0aa12fb
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433238"
---
# <a name="office"></a><span data-ttu-id="702dd-102">Office</span><span class="sxs-lookup"><span data-stu-id="702dd-102">Office</span></span>

<span data-ttu-id="702dd-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="702dd-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="702dd-105">Требования</span><span class="sxs-lookup"><span data-stu-id="702dd-105">Requirements</span></span>

|<span data-ttu-id="702dd-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="702dd-106">Requirement</span></span>| <span data-ttu-id="702dd-107">Значение</span><span class="sxs-lookup"><span data-stu-id="702dd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="702dd-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="702dd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="702dd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="702dd-109">1.0</span></span>|
|[<span data-ttu-id="702dd-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="702dd-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="702dd-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="702dd-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="702dd-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="702dd-112">Namespaces</span></span>

<span data-ttu-id="702dd-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="702dd-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="702dd-114">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="702dd-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="702dd-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="702dd-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="702dd-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="702dd-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="702dd-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="702dd-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="702dd-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="702dd-118">Type:</span></span>

*   <span data-ttu-id="702dd-119">String</span><span class="sxs-lookup"><span data-stu-id="702dd-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="702dd-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="702dd-120">Properties:</span></span>

|<span data-ttu-id="702dd-121">Имя</span><span class="sxs-lookup"><span data-stu-id="702dd-121">Name</span></span>| <span data-ttu-id="702dd-122">Тип</span><span class="sxs-lookup"><span data-stu-id="702dd-122">Type</span></span>| <span data-ttu-id="702dd-123">Описание</span><span class="sxs-lookup"><span data-stu-id="702dd-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="702dd-124">Для указания</span><span class="sxs-lookup"><span data-stu-id="702dd-124">String</span></span>|<span data-ttu-id="702dd-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="702dd-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="702dd-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="702dd-126">String</span></span>|<span data-ttu-id="702dd-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="702dd-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="702dd-128">Требования</span><span class="sxs-lookup"><span data-stu-id="702dd-128">Requirements</span></span>

|<span data-ttu-id="702dd-129">Requirement</span><span class="sxs-lookup"><span data-stu-id="702dd-129">Requirement</span></span>| <span data-ttu-id="702dd-130">Значение</span><span class="sxs-lookup"><span data-stu-id="702dd-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="702dd-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="702dd-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="702dd-132">1.0</span><span class="sxs-lookup"><span data-stu-id="702dd-132">1.0</span></span>|
|[<span data-ttu-id="702dd-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="702dd-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="702dd-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="702dd-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="702dd-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="702dd-135">CoercionType :String</span></span>

<span data-ttu-id="702dd-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="702dd-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="702dd-137">Тип:</span><span class="sxs-lookup"><span data-stu-id="702dd-137">Type:</span></span>

*   <span data-ttu-id="702dd-138">String</span><span class="sxs-lookup"><span data-stu-id="702dd-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="702dd-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="702dd-139">Properties:</span></span>

|<span data-ttu-id="702dd-140">Имя</span><span class="sxs-lookup"><span data-stu-id="702dd-140">Name</span></span>| <span data-ttu-id="702dd-141">Тип</span><span class="sxs-lookup"><span data-stu-id="702dd-141">Type</span></span>| <span data-ttu-id="702dd-142">Описание</span><span class="sxs-lookup"><span data-stu-id="702dd-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="702dd-143">String</span><span class="sxs-lookup"><span data-stu-id="702dd-143">String</span></span>|<span data-ttu-id="702dd-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="702dd-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="702dd-145">String</span><span class="sxs-lookup"><span data-stu-id="702dd-145">String</span></span>|<span data-ttu-id="702dd-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="702dd-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="702dd-147">Требования</span><span class="sxs-lookup"><span data-stu-id="702dd-147">Requirements</span></span>

|<span data-ttu-id="702dd-148">Requirement</span><span class="sxs-lookup"><span data-stu-id="702dd-148">Requirement</span></span>| <span data-ttu-id="702dd-149">Значение</span><span class="sxs-lookup"><span data-stu-id="702dd-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="702dd-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="702dd-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="702dd-151">1.0</span><span class="sxs-lookup"><span data-stu-id="702dd-151">1.0</span></span>|
|[<span data-ttu-id="702dd-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="702dd-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="702dd-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="702dd-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="702dd-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="702dd-154">SourceProperty :String</span></span>

<span data-ttu-id="702dd-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="702dd-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="702dd-156">Тип:</span><span class="sxs-lookup"><span data-stu-id="702dd-156">Type:</span></span>

*   <span data-ttu-id="702dd-157">String</span><span class="sxs-lookup"><span data-stu-id="702dd-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="702dd-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="702dd-158">Properties:</span></span>

|<span data-ttu-id="702dd-159">Имя</span><span class="sxs-lookup"><span data-stu-id="702dd-159">Name</span></span>| <span data-ttu-id="702dd-160">Тип</span><span class="sxs-lookup"><span data-stu-id="702dd-160">Type</span></span>| <span data-ttu-id="702dd-161">Описание</span><span class="sxs-lookup"><span data-stu-id="702dd-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="702dd-162">String</span><span class="sxs-lookup"><span data-stu-id="702dd-162">String</span></span>|<span data-ttu-id="702dd-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="702dd-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="702dd-164">String</span><span class="sxs-lookup"><span data-stu-id="702dd-164">String</span></span>|<span data-ttu-id="702dd-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="702dd-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="702dd-166">Требования</span><span class="sxs-lookup"><span data-stu-id="702dd-166">Requirements</span></span>

|<span data-ttu-id="702dd-167">Requirement</span><span class="sxs-lookup"><span data-stu-id="702dd-167">Requirement</span></span>| <span data-ttu-id="702dd-168">Значение</span><span class="sxs-lookup"><span data-stu-id="702dd-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="702dd-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="702dd-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="702dd-170">1.0</span><span class="sxs-lookup"><span data-stu-id="702dd-170">1.0</span></span>|
|[<span data-ttu-id="702dd-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="702dd-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="702dd-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="702dd-172">Compose or read</span></span>|