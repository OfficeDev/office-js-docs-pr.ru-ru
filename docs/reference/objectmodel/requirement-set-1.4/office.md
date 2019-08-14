---
title: Пространство имен Office — набор обязательных элементов 1,4
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 33089ad278613dac76a8610aba803c2f9db1d088
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395682"
---
# <a name="office"></a><span data-ttu-id="243ca-102">Office</span><span class="sxs-lookup"><span data-stu-id="243ca-102">Office</span></span>

<span data-ttu-id="243ca-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="243ca-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="243ca-105">Требования</span><span class="sxs-lookup"><span data-stu-id="243ca-105">Requirements</span></span>

|<span data-ttu-id="243ca-106">Требование</span><span class="sxs-lookup"><span data-stu-id="243ca-106">Requirement</span></span>| <span data-ttu-id="243ca-107">Значение</span><span class="sxs-lookup"><span data-stu-id="243ca-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="243ca-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="243ca-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243ca-109">1.0</span><span class="sxs-lookup"><span data-stu-id="243ca-109">1.0</span></span>|
|[<span data-ttu-id="243ca-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="243ca-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="243ca-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="243ca-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="243ca-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="243ca-112">Members and methods</span></span>

| <span data-ttu-id="243ca-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="243ca-113">Member</span></span> | <span data-ttu-id="243ca-114">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="243ca-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="243ca-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="243ca-116">Member</span><span class="sxs-lookup"><span data-stu-id="243ca-116">Member</span></span> |
| [<span data-ttu-id="243ca-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="243ca-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="243ca-118">Member</span><span class="sxs-lookup"><span data-stu-id="243ca-118">Member</span></span> |
| [<span data-ttu-id="243ca-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="243ca-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="243ca-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="243ca-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="243ca-121">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="243ca-121">Namespaces</span></span>

<span data-ttu-id="243ca-122">[context.](Office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="243ca-122">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="243ca-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="243ca-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="243ca-124">Members</span><span class="sxs-lookup"><span data-stu-id="243ca-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="243ca-125">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="243ca-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="243ca-126">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="243ca-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="243ca-127">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-127">Type</span></span>

*   <span data-ttu-id="243ca-128">String</span><span class="sxs-lookup"><span data-stu-id="243ca-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243ca-129">Свойства:</span><span class="sxs-lookup"><span data-stu-id="243ca-129">Properties:</span></span>

|<span data-ttu-id="243ca-130">Имя</span><span class="sxs-lookup"><span data-stu-id="243ca-130">Name</span></span>| <span data-ttu-id="243ca-131">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-131">Type</span></span>| <span data-ttu-id="243ca-132">Описание</span><span class="sxs-lookup"><span data-stu-id="243ca-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="243ca-133">String</span><span class="sxs-lookup"><span data-stu-id="243ca-133">String</span></span>|<span data-ttu-id="243ca-134">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="243ca-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="243ca-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="243ca-135">String</span></span>|<span data-ttu-id="243ca-136">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="243ca-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243ca-137">Требования</span><span class="sxs-lookup"><span data-stu-id="243ca-137">Requirements</span></span>

|<span data-ttu-id="243ca-138">Требование</span><span class="sxs-lookup"><span data-stu-id="243ca-138">Requirement</span></span>| <span data-ttu-id="243ca-139">Значение</span><span class="sxs-lookup"><span data-stu-id="243ca-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="243ca-140">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="243ca-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243ca-141">1.0</span><span class="sxs-lookup"><span data-stu-id="243ca-141">1.0</span></span>|
|[<span data-ttu-id="243ca-142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="243ca-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="243ca-143">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="243ca-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="243ca-144">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="243ca-144">CoercionType: String</span></span>

<span data-ttu-id="243ca-145">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="243ca-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="243ca-146">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-146">Type</span></span>

*   <span data-ttu-id="243ca-147">String</span><span class="sxs-lookup"><span data-stu-id="243ca-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243ca-148">Свойства:</span><span class="sxs-lookup"><span data-stu-id="243ca-148">Properties:</span></span>

|<span data-ttu-id="243ca-149">Имя</span><span class="sxs-lookup"><span data-stu-id="243ca-149">Name</span></span>| <span data-ttu-id="243ca-150">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-150">Type</span></span>| <span data-ttu-id="243ca-151">Описание</span><span class="sxs-lookup"><span data-stu-id="243ca-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="243ca-152">String</span><span class="sxs-lookup"><span data-stu-id="243ca-152">String</span></span>|<span data-ttu-id="243ca-153">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="243ca-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="243ca-154">String</span><span class="sxs-lookup"><span data-stu-id="243ca-154">String</span></span>|<span data-ttu-id="243ca-155">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="243ca-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243ca-156">Требования</span><span class="sxs-lookup"><span data-stu-id="243ca-156">Requirements</span></span>

|<span data-ttu-id="243ca-157">Требование</span><span class="sxs-lookup"><span data-stu-id="243ca-157">Requirement</span></span>| <span data-ttu-id="243ca-158">Значение</span><span class="sxs-lookup"><span data-stu-id="243ca-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="243ca-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="243ca-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243ca-160">1.0</span><span class="sxs-lookup"><span data-stu-id="243ca-160">1.0</span></span>|
|[<span data-ttu-id="243ca-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="243ca-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="243ca-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="243ca-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="243ca-163">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="243ca-163">SourceProperty: String</span></span>

<span data-ttu-id="243ca-164">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="243ca-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="243ca-165">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-165">Type</span></span>

*   <span data-ttu-id="243ca-166">String</span><span class="sxs-lookup"><span data-stu-id="243ca-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243ca-167">Свойства:</span><span class="sxs-lookup"><span data-stu-id="243ca-167">Properties:</span></span>

|<span data-ttu-id="243ca-168">Имя</span><span class="sxs-lookup"><span data-stu-id="243ca-168">Name</span></span>| <span data-ttu-id="243ca-169">Тип</span><span class="sxs-lookup"><span data-stu-id="243ca-169">Type</span></span>| <span data-ttu-id="243ca-170">Описание</span><span class="sxs-lookup"><span data-stu-id="243ca-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="243ca-171">String</span><span class="sxs-lookup"><span data-stu-id="243ca-171">String</span></span>|<span data-ttu-id="243ca-172">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="243ca-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="243ca-173">String</span><span class="sxs-lookup"><span data-stu-id="243ca-173">String</span></span>|<span data-ttu-id="243ca-174">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="243ca-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243ca-175">Требования</span><span class="sxs-lookup"><span data-stu-id="243ca-175">Requirements</span></span>

|<span data-ttu-id="243ca-176">Требование</span><span class="sxs-lookup"><span data-stu-id="243ca-176">Requirement</span></span>| <span data-ttu-id="243ca-177">Значение</span><span class="sxs-lookup"><span data-stu-id="243ca-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="243ca-178">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="243ca-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243ca-179">1.0</span><span class="sxs-lookup"><span data-stu-id="243ca-179">1.0</span></span>|
|[<span data-ttu-id="243ca-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="243ca-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="243ca-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="243ca-181">Compose or Read</span></span>|
