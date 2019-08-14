---
title: Пространство имен Office — набор обязательных элементов 1,2
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 5a8431580fce2a98f2076ef3df151f08d5435d54
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395647"
---
# <a name="office"></a><span data-ttu-id="419b0-102">Office</span><span class="sxs-lookup"><span data-stu-id="419b0-102">Office</span></span>

<span data-ttu-id="419b0-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="419b0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="419b0-105">Требования</span><span class="sxs-lookup"><span data-stu-id="419b0-105">Requirements</span></span>

|<span data-ttu-id="419b0-106">Требование</span><span class="sxs-lookup"><span data-stu-id="419b0-106">Requirement</span></span>| <span data-ttu-id="419b0-107">Значение</span><span class="sxs-lookup"><span data-stu-id="419b0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="419b0-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="419b0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="419b0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="419b0-109">1.0</span></span>|
|[<span data-ttu-id="419b0-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="419b0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="419b0-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="419b0-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="419b0-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="419b0-112">Members and methods</span></span>

| <span data-ttu-id="419b0-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="419b0-113">Member</span></span> | <span data-ttu-id="419b0-114">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="419b0-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="419b0-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="419b0-116">Member</span><span class="sxs-lookup"><span data-stu-id="419b0-116">Member</span></span> |
| [<span data-ttu-id="419b0-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="419b0-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="419b0-118">Member</span><span class="sxs-lookup"><span data-stu-id="419b0-118">Member</span></span> |
| [<span data-ttu-id="419b0-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="419b0-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="419b0-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="419b0-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="419b0-121">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="419b0-121">Namespaces</span></span>

<span data-ttu-id="419b0-122">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="419b0-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="419b0-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="419b0-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="419b0-124">Members</span><span class="sxs-lookup"><span data-stu-id="419b0-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="419b0-125">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="419b0-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="419b0-126">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="419b0-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="419b0-127">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-127">Type</span></span>

*   <span data-ttu-id="419b0-128">String</span><span class="sxs-lookup"><span data-stu-id="419b0-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="419b0-129">Свойства:</span><span class="sxs-lookup"><span data-stu-id="419b0-129">Properties:</span></span>

|<span data-ttu-id="419b0-130">Имя</span><span class="sxs-lookup"><span data-stu-id="419b0-130">Name</span></span>| <span data-ttu-id="419b0-131">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-131">Type</span></span>| <span data-ttu-id="419b0-132">Описание</span><span class="sxs-lookup"><span data-stu-id="419b0-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="419b0-133">String</span><span class="sxs-lookup"><span data-stu-id="419b0-133">String</span></span>|<span data-ttu-id="419b0-134">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="419b0-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="419b0-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="419b0-135">String</span></span>|<span data-ttu-id="419b0-136">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="419b0-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="419b0-137">Требования</span><span class="sxs-lookup"><span data-stu-id="419b0-137">Requirements</span></span>

|<span data-ttu-id="419b0-138">Требование</span><span class="sxs-lookup"><span data-stu-id="419b0-138">Requirement</span></span>| <span data-ttu-id="419b0-139">Значение</span><span class="sxs-lookup"><span data-stu-id="419b0-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="419b0-140">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="419b0-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="419b0-141">1.0</span><span class="sxs-lookup"><span data-stu-id="419b0-141">1.0</span></span>|
|[<span data-ttu-id="419b0-142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="419b0-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="419b0-143">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="419b0-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="419b0-144">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="419b0-144">CoercionType: String</span></span>

<span data-ttu-id="419b0-145">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="419b0-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="419b0-146">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-146">Type</span></span>

*   <span data-ttu-id="419b0-147">String</span><span class="sxs-lookup"><span data-stu-id="419b0-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="419b0-148">Свойства:</span><span class="sxs-lookup"><span data-stu-id="419b0-148">Properties:</span></span>

|<span data-ttu-id="419b0-149">Имя</span><span class="sxs-lookup"><span data-stu-id="419b0-149">Name</span></span>| <span data-ttu-id="419b0-150">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-150">Type</span></span>| <span data-ttu-id="419b0-151">Описание</span><span class="sxs-lookup"><span data-stu-id="419b0-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="419b0-152">String</span><span class="sxs-lookup"><span data-stu-id="419b0-152">String</span></span>|<span data-ttu-id="419b0-153">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="419b0-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="419b0-154">String</span><span class="sxs-lookup"><span data-stu-id="419b0-154">String</span></span>|<span data-ttu-id="419b0-155">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="419b0-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="419b0-156">Требования</span><span class="sxs-lookup"><span data-stu-id="419b0-156">Requirements</span></span>

|<span data-ttu-id="419b0-157">Требование</span><span class="sxs-lookup"><span data-stu-id="419b0-157">Requirement</span></span>| <span data-ttu-id="419b0-158">Значение</span><span class="sxs-lookup"><span data-stu-id="419b0-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="419b0-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="419b0-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="419b0-160">1.0</span><span class="sxs-lookup"><span data-stu-id="419b0-160">1.0</span></span>|
|[<span data-ttu-id="419b0-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="419b0-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="419b0-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="419b0-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="419b0-163">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="419b0-163">SourceProperty: String</span></span>

<span data-ttu-id="419b0-164">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="419b0-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="419b0-165">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-165">Type</span></span>

*   <span data-ttu-id="419b0-166">String</span><span class="sxs-lookup"><span data-stu-id="419b0-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="419b0-167">Свойства:</span><span class="sxs-lookup"><span data-stu-id="419b0-167">Properties:</span></span>

|<span data-ttu-id="419b0-168">Имя</span><span class="sxs-lookup"><span data-stu-id="419b0-168">Name</span></span>| <span data-ttu-id="419b0-169">Тип</span><span class="sxs-lookup"><span data-stu-id="419b0-169">Type</span></span>| <span data-ttu-id="419b0-170">Описание</span><span class="sxs-lookup"><span data-stu-id="419b0-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="419b0-171">String</span><span class="sxs-lookup"><span data-stu-id="419b0-171">String</span></span>|<span data-ttu-id="419b0-172">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="419b0-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="419b0-173">String</span><span class="sxs-lookup"><span data-stu-id="419b0-173">String</span></span>|<span data-ttu-id="419b0-174">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="419b0-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="419b0-175">Требования</span><span class="sxs-lookup"><span data-stu-id="419b0-175">Requirements</span></span>

|<span data-ttu-id="419b0-176">Требование</span><span class="sxs-lookup"><span data-stu-id="419b0-176">Requirement</span></span>| <span data-ttu-id="419b0-177">Значение</span><span class="sxs-lookup"><span data-stu-id="419b0-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="419b0-178">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="419b0-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="419b0-179">1.0</span><span class="sxs-lookup"><span data-stu-id="419b0-179">1.0</span></span>|
|[<span data-ttu-id="419b0-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="419b0-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="419b0-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="419b0-181">Compose or Read</span></span>|
