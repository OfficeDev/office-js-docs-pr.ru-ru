---
title: Пространство имен Office — набор обязательных элементов 1,1
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: b22dbae0824c8b5047ce90c255e06f09744dd05c
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890762"
---
# <a name="office-mailbox-requirement-set-11"></a><span data-ttu-id="af73f-103">Office (набор требований для почтового ящика 1,1)</span><span class="sxs-lookup"><span data-stu-id="af73f-103">Office (Mailbox requirement set 1.1)</span></span>

<span data-ttu-id="af73f-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="af73f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="af73f-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="af73f-106">Requirements</span></span>

|<span data-ttu-id="af73f-107">Требование</span><span class="sxs-lookup"><span data-stu-id="af73f-107">Requirement</span></span>| <span data-ttu-id="af73f-108">Значение</span><span class="sxs-lookup"><span data-stu-id="af73f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="af73f-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="af73f-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af73f-110">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-110">1.1</span></span>|
|[<span data-ttu-id="af73f-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="af73f-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af73f-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="af73f-113">Properties</span><span class="sxs-lookup"><span data-stu-id="af73f-113">Properties</span></span>

| <span data-ttu-id="af73f-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="af73f-114">Property</span></span> | <span data-ttu-id="af73f-115">Способов</span><span class="sxs-lookup"><span data-stu-id="af73f-115">Modes</span></span> | <span data-ttu-id="af73f-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="af73f-116">Return type</span></span> | <span data-ttu-id="af73f-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="af73f-117">Minimum</span></span><br><span data-ttu-id="af73f-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="af73f-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="af73f-119">контекст</span><span class="sxs-lookup"><span data-stu-id="af73f-119">context</span></span>](office.context.md) | <span data-ttu-id="af73f-120">Создание</span><span class="sxs-lookup"><span data-stu-id="af73f-120">Compose</span></span><br><span data-ttu-id="af73f-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-121">Read</span></span> | [<span data-ttu-id="af73f-122">Context</span><span class="sxs-lookup"><span data-stu-id="af73f-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="af73f-123">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="af73f-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="af73f-124">Enumerations</span></span>

| <span data-ttu-id="af73f-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="af73f-125">Enumeration</span></span> | <span data-ttu-id="af73f-126">Способов</span><span class="sxs-lookup"><span data-stu-id="af73f-126">Modes</span></span> | <span data-ttu-id="af73f-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="af73f-127">Return type</span></span> | <span data-ttu-id="af73f-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="af73f-128">Minimum</span></span><br><span data-ttu-id="af73f-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="af73f-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="af73f-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="af73f-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="af73f-131">Создание</span><span class="sxs-lookup"><span data-stu-id="af73f-131">Compose</span></span><br><span data-ttu-id="af73f-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-132">Read</span></span> | <span data-ttu-id="af73f-133">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-133">String</span></span> | [<span data-ttu-id="af73f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af73f-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="af73f-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="af73f-136">Создание</span><span class="sxs-lookup"><span data-stu-id="af73f-136">Compose</span></span><br><span data-ttu-id="af73f-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-137">Read</span></span> | <span data-ttu-id="af73f-138">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-138">String</span></span> | [<span data-ttu-id="af73f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af73f-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="af73f-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="af73f-141">Создание</span><span class="sxs-lookup"><span data-stu-id="af73f-141">Compose</span></span><br><span data-ttu-id="af73f-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-142">Read</span></span> | <span data-ttu-id="af73f-143">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-143">String</span></span> | [<span data-ttu-id="af73f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="af73f-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="af73f-145">Namespaces</span></span>

<span data-ttu-id="af73f-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="af73f-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="af73f-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="af73f-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="af73f-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="af73f-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="af73f-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="af73f-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="af73f-150">Тип</span><span class="sxs-lookup"><span data-stu-id="af73f-150">Type</span></span>

*   <span data-ttu-id="af73f-151">String</span><span class="sxs-lookup"><span data-stu-id="af73f-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af73f-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="af73f-152">Properties:</span></span>

|<span data-ttu-id="af73f-153">Имя</span><span class="sxs-lookup"><span data-stu-id="af73f-153">Name</span></span>| <span data-ttu-id="af73f-154">Тип</span><span class="sxs-lookup"><span data-stu-id="af73f-154">Type</span></span>| <span data-ttu-id="af73f-155">Описание</span><span class="sxs-lookup"><span data-stu-id="af73f-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="af73f-156">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-156">String</span></span>|<span data-ttu-id="af73f-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="af73f-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="af73f-158">Для указания</span><span class="sxs-lookup"><span data-stu-id="af73f-158">String</span></span>|<span data-ttu-id="af73f-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="af73f-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af73f-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="af73f-160">Requirements</span></span>

|<span data-ttu-id="af73f-161">Требование</span><span class="sxs-lookup"><span data-stu-id="af73f-161">Requirement</span></span>| <span data-ttu-id="af73f-162">Значение</span><span class="sxs-lookup"><span data-stu-id="af73f-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="af73f-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="af73f-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af73f-164">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-164">1.1</span></span>|
|[<span data-ttu-id="af73f-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="af73f-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af73f-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="af73f-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="af73f-167">CoercionType: String</span></span>

<span data-ttu-id="af73f-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="af73f-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="af73f-169">Тип</span><span class="sxs-lookup"><span data-stu-id="af73f-169">Type</span></span>

*   <span data-ttu-id="af73f-170">String</span><span class="sxs-lookup"><span data-stu-id="af73f-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af73f-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="af73f-171">Properties:</span></span>

|<span data-ttu-id="af73f-172">Имя</span><span class="sxs-lookup"><span data-stu-id="af73f-172">Name</span></span>| <span data-ttu-id="af73f-173">Тип</span><span class="sxs-lookup"><span data-stu-id="af73f-173">Type</span></span>| <span data-ttu-id="af73f-174">Описание</span><span class="sxs-lookup"><span data-stu-id="af73f-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="af73f-175">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-175">String</span></span>|<span data-ttu-id="af73f-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="af73f-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="af73f-177">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-177">String</span></span>|<span data-ttu-id="af73f-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="af73f-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af73f-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="af73f-179">Requirements</span></span>

|<span data-ttu-id="af73f-180">Требование</span><span class="sxs-lookup"><span data-stu-id="af73f-180">Requirement</span></span>| <span data-ttu-id="af73f-181">Значение</span><span class="sxs-lookup"><span data-stu-id="af73f-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="af73f-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="af73f-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af73f-183">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-183">1.1</span></span>|
|[<span data-ttu-id="af73f-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="af73f-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af73f-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="af73f-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="af73f-186">SourceProperty: String</span></span>

<span data-ttu-id="af73f-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="af73f-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="af73f-188">Тип</span><span class="sxs-lookup"><span data-stu-id="af73f-188">Type</span></span>

*   <span data-ttu-id="af73f-189">String</span><span class="sxs-lookup"><span data-stu-id="af73f-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af73f-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="af73f-190">Properties:</span></span>

|<span data-ttu-id="af73f-191">Имя</span><span class="sxs-lookup"><span data-stu-id="af73f-191">Name</span></span>| <span data-ttu-id="af73f-192">Тип</span><span class="sxs-lookup"><span data-stu-id="af73f-192">Type</span></span>| <span data-ttu-id="af73f-193">Описание</span><span class="sxs-lookup"><span data-stu-id="af73f-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="af73f-194">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-194">String</span></span>|<span data-ttu-id="af73f-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="af73f-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="af73f-196">Строка</span><span class="sxs-lookup"><span data-stu-id="af73f-196">String</span></span>|<span data-ttu-id="af73f-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="af73f-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af73f-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="af73f-198">Requirements</span></span>

|<span data-ttu-id="af73f-199">Требование</span><span class="sxs-lookup"><span data-stu-id="af73f-199">Requirement</span></span>| <span data-ttu-id="af73f-200">Значение</span><span class="sxs-lookup"><span data-stu-id="af73f-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="af73f-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="af73f-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af73f-202">1.1</span><span class="sxs-lookup"><span data-stu-id="af73f-202">1.1</span></span>|
|[<span data-ttu-id="af73f-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="af73f-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af73f-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="af73f-204">Compose or Read</span></span>|