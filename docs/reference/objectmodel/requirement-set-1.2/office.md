---
title: Пространство имен Office — набор обязательных элементов 1.2
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: dd623959c7c71f6bb7f837e73b1713be41639de5
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457791"
---
# <a name="office"></a><span data-ttu-id="53f7a-102">Office</span><span class="sxs-lookup"><span data-stu-id="53f7a-102">Office</span></span>

<span data-ttu-id="53f7a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="53f7a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="53f7a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="53f7a-105">Requirements</span></span>

|<span data-ttu-id="53f7a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="53f7a-106">Requirement</span></span>| <span data-ttu-id="53f7a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="53f7a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="53f7a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53f7a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53f7a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="53f7a-109">1.0</span></span>|
|[<span data-ttu-id="53f7a-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53f7a-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53f7a-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53f7a-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="53f7a-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="53f7a-112">Namespaces</span></span>

<span data-ttu-id="53f7a-113">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="53f7a-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="53f7a-114">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="53f7a-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="53f7a-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="53f7a-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="53f7a-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="53f7a-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="53f7a-117">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="53f7a-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="53f7a-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="53f7a-118">Type:</span></span>

*   <span data-ttu-id="53f7a-119">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53f7a-120">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53f7a-120">Properties:</span></span>

|<span data-ttu-id="53f7a-121">Имя</span><span class="sxs-lookup"><span data-stu-id="53f7a-121">Name</span></span>| <span data-ttu-id="53f7a-122">Тип</span><span class="sxs-lookup"><span data-stu-id="53f7a-122">Type</span></span>| <span data-ttu-id="53f7a-123">Описание</span><span class="sxs-lookup"><span data-stu-id="53f7a-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="53f7a-124">Для указания</span><span class="sxs-lookup"><span data-stu-id="53f7a-124">String</span></span>|<span data-ttu-id="53f7a-125">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="53f7a-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="53f7a-126">Для указания</span><span class="sxs-lookup"><span data-stu-id="53f7a-126">String</span></span>|<span data-ttu-id="53f7a-127">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="53f7a-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53f7a-128">Требования</span><span class="sxs-lookup"><span data-stu-id="53f7a-128">Requirements</span></span>

|<span data-ttu-id="53f7a-129">Требование</span><span class="sxs-lookup"><span data-stu-id="53f7a-129">Requirement</span></span>| <span data-ttu-id="53f7a-130">Значение</span><span class="sxs-lookup"><span data-stu-id="53f7a-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="53f7a-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53f7a-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53f7a-132">1.0</span><span class="sxs-lookup"><span data-stu-id="53f7a-132">1.0</span></span>|
|[<span data-ttu-id="53f7a-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53f7a-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53f7a-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53f7a-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="53f7a-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="53f7a-135">CoercionType :String</span></span>

<span data-ttu-id="53f7a-136">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="53f7a-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="53f7a-137">Тип:</span><span class="sxs-lookup"><span data-stu-id="53f7a-137">Type:</span></span>

*   <span data-ttu-id="53f7a-138">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53f7a-139">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53f7a-139">Properties:</span></span>

|<span data-ttu-id="53f7a-140">Имя</span><span class="sxs-lookup"><span data-stu-id="53f7a-140">Name</span></span>| <span data-ttu-id="53f7a-141">Тип</span><span class="sxs-lookup"><span data-stu-id="53f7a-141">Type</span></span>| <span data-ttu-id="53f7a-142">Описание</span><span class="sxs-lookup"><span data-stu-id="53f7a-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="53f7a-143">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-143">String</span></span>|<span data-ttu-id="53f7a-144">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="53f7a-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="53f7a-145">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-145">String</span></span>|<span data-ttu-id="53f7a-146">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="53f7a-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53f7a-147">Требования</span><span class="sxs-lookup"><span data-stu-id="53f7a-147">Requirements</span></span>

|<span data-ttu-id="53f7a-148">Требование</span><span class="sxs-lookup"><span data-stu-id="53f7a-148">Requirement</span></span>| <span data-ttu-id="53f7a-149">Значение</span><span class="sxs-lookup"><span data-stu-id="53f7a-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="53f7a-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53f7a-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53f7a-151">1.0</span><span class="sxs-lookup"><span data-stu-id="53f7a-151">1.0</span></span>|
|[<span data-ttu-id="53f7a-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53f7a-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53f7a-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53f7a-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="53f7a-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="53f7a-154">SourceProperty :String</span></span>

<span data-ttu-id="53f7a-155">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="53f7a-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="53f7a-156">Тип:</span><span class="sxs-lookup"><span data-stu-id="53f7a-156">Type:</span></span>

*   <span data-ttu-id="53f7a-157">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53f7a-158">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53f7a-158">Properties:</span></span>

|<span data-ttu-id="53f7a-159">Имя</span><span class="sxs-lookup"><span data-stu-id="53f7a-159">Name</span></span>| <span data-ttu-id="53f7a-160">Тип</span><span class="sxs-lookup"><span data-stu-id="53f7a-160">Type</span></span>| <span data-ttu-id="53f7a-161">Описание</span><span class="sxs-lookup"><span data-stu-id="53f7a-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="53f7a-162">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-162">String</span></span>|<span data-ttu-id="53f7a-163">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="53f7a-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="53f7a-164">String</span><span class="sxs-lookup"><span data-stu-id="53f7a-164">String</span></span>|<span data-ttu-id="53f7a-165">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="53f7a-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53f7a-166">Требования</span><span class="sxs-lookup"><span data-stu-id="53f7a-166">Requirements</span></span>

|<span data-ttu-id="53f7a-167">Требование</span><span class="sxs-lookup"><span data-stu-id="53f7a-167">Requirement</span></span>| <span data-ttu-id="53f7a-168">Значение</span><span class="sxs-lookup"><span data-stu-id="53f7a-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="53f7a-169">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53f7a-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53f7a-170">1.0</span><span class="sxs-lookup"><span data-stu-id="53f7a-170">1.0</span></span>|
|[<span data-ttu-id="53f7a-171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53f7a-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53f7a-172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53f7a-172">Compose or read</span></span>|