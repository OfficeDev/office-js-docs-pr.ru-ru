---
title: Пространство имен Office — набор обязательных элементов 1.3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 0b22574693fb129be6a08a89b58beceb746fa283
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268399"
---
# <a name="office"></a><span data-ttu-id="1614a-102">Office</span><span class="sxs-lookup"><span data-stu-id="1614a-102">Office</span></span>

<span data-ttu-id="1614a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1614a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1614a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="1614a-105">Requirements</span></span>

|<span data-ttu-id="1614a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="1614a-106">Requirement</span></span>| <span data-ttu-id="1614a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="1614a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1614a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1614a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1614a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1614a-109">1.0</span></span>|
|[<span data-ttu-id="1614a-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1614a-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1614a-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1614a-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1614a-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="1614a-112">Members and methods</span></span>

| <span data-ttu-id="1614a-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="1614a-113">Member</span></span> | <span data-ttu-id="1614a-114">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1614a-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1614a-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1614a-116">Member</span><span class="sxs-lookup"><span data-stu-id="1614a-116">Member</span></span> |
| [<span data-ttu-id="1614a-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1614a-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1614a-118">Member</span><span class="sxs-lookup"><span data-stu-id="1614a-118">Member</span></span> |
| [<span data-ttu-id="1614a-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1614a-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1614a-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="1614a-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1614a-121">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="1614a-121">Namespaces</span></span>

<span data-ttu-id="1614a-122">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="1614a-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="1614a-123">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="1614a-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="1614a-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="1614a-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1614a-125">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="1614a-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="1614a-126">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="1614a-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1614a-127">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-127">Type</span></span>

*   <span data-ttu-id="1614a-128">String</span><span class="sxs-lookup"><span data-stu-id="1614a-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1614a-129">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1614a-129">Properties:</span></span>

|<span data-ttu-id="1614a-130">Имя</span><span class="sxs-lookup"><span data-stu-id="1614a-130">Name</span></span>| <span data-ttu-id="1614a-131">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-131">Type</span></span>| <span data-ttu-id="1614a-132">Описание</span><span class="sxs-lookup"><span data-stu-id="1614a-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1614a-133">String</span><span class="sxs-lookup"><span data-stu-id="1614a-133">String</span></span>|<span data-ttu-id="1614a-134">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="1614a-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1614a-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="1614a-135">String</span></span>|<span data-ttu-id="1614a-136">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="1614a-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1614a-137">Требования</span><span class="sxs-lookup"><span data-stu-id="1614a-137">Requirements</span></span>

|<span data-ttu-id="1614a-138">Требование</span><span class="sxs-lookup"><span data-stu-id="1614a-138">Requirement</span></span>| <span data-ttu-id="1614a-139">Значение</span><span class="sxs-lookup"><span data-stu-id="1614a-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="1614a-140">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1614a-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1614a-141">1.0</span><span class="sxs-lookup"><span data-stu-id="1614a-141">1.0</span></span>|
|[<span data-ttu-id="1614a-142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1614a-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1614a-143">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1614a-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="1614a-144">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="1614a-144">CoercionType: String</span></span>

<span data-ttu-id="1614a-145">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="1614a-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1614a-146">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-146">Type</span></span>

*   <span data-ttu-id="1614a-147">String</span><span class="sxs-lookup"><span data-stu-id="1614a-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1614a-148">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1614a-148">Properties:</span></span>

|<span data-ttu-id="1614a-149">Имя</span><span class="sxs-lookup"><span data-stu-id="1614a-149">Name</span></span>| <span data-ttu-id="1614a-150">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-150">Type</span></span>| <span data-ttu-id="1614a-151">Описание</span><span class="sxs-lookup"><span data-stu-id="1614a-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1614a-152">String</span><span class="sxs-lookup"><span data-stu-id="1614a-152">String</span></span>|<span data-ttu-id="1614a-153">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="1614a-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1614a-154">String</span><span class="sxs-lookup"><span data-stu-id="1614a-154">String</span></span>|<span data-ttu-id="1614a-155">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="1614a-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1614a-156">Требования</span><span class="sxs-lookup"><span data-stu-id="1614a-156">Requirements</span></span>

|<span data-ttu-id="1614a-157">Требование</span><span class="sxs-lookup"><span data-stu-id="1614a-157">Requirement</span></span>| <span data-ttu-id="1614a-158">Значение</span><span class="sxs-lookup"><span data-stu-id="1614a-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="1614a-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1614a-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1614a-160">1.0</span><span class="sxs-lookup"><span data-stu-id="1614a-160">1.0</span></span>|
|[<span data-ttu-id="1614a-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1614a-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1614a-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1614a-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="1614a-163">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="1614a-163">SourceProperty: String</span></span>

<span data-ttu-id="1614a-164">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="1614a-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1614a-165">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-165">Type</span></span>

*   <span data-ttu-id="1614a-166">String</span><span class="sxs-lookup"><span data-stu-id="1614a-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1614a-167">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1614a-167">Properties:</span></span>

|<span data-ttu-id="1614a-168">Имя</span><span class="sxs-lookup"><span data-stu-id="1614a-168">Name</span></span>| <span data-ttu-id="1614a-169">Тип</span><span class="sxs-lookup"><span data-stu-id="1614a-169">Type</span></span>| <span data-ttu-id="1614a-170">Описание</span><span class="sxs-lookup"><span data-stu-id="1614a-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1614a-171">String</span><span class="sxs-lookup"><span data-stu-id="1614a-171">String</span></span>|<span data-ttu-id="1614a-172">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="1614a-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1614a-173">String</span><span class="sxs-lookup"><span data-stu-id="1614a-173">String</span></span>|<span data-ttu-id="1614a-174">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="1614a-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1614a-175">Требования</span><span class="sxs-lookup"><span data-stu-id="1614a-175">Requirements</span></span>

|<span data-ttu-id="1614a-176">Требование</span><span class="sxs-lookup"><span data-stu-id="1614a-176">Requirement</span></span>| <span data-ttu-id="1614a-177">Значение</span><span class="sxs-lookup"><span data-stu-id="1614a-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="1614a-178">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1614a-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1614a-179">1.0</span><span class="sxs-lookup"><span data-stu-id="1614a-179">1.0</span></span>|
|[<span data-ttu-id="1614a-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1614a-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1614a-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1614a-181">Compose or Read</span></span>|
