---
title: Пространство имен Office — набор обязательных элементов 1.3
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,3.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: eb3ebba75df8345402ab0ce4ac2b5cc5f0354e6c
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890699"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="e7efd-103">Office (набор требований для почтового ящика 1,3)</span><span class="sxs-lookup"><span data-stu-id="e7efd-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="e7efd-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e7efd-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7efd-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7efd-106">Requirements</span></span>

|<span data-ttu-id="e7efd-107">Требование</span><span class="sxs-lookup"><span data-stu-id="e7efd-107">Requirement</span></span>| <span data-ttu-id="e7efd-108">Значение</span><span class="sxs-lookup"><span data-stu-id="e7efd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7efd-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7efd-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e7efd-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-110">1.1</span></span>|
|[<span data-ttu-id="e7efd-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7efd-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e7efd-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e7efd-113">Properties</span><span class="sxs-lookup"><span data-stu-id="e7efd-113">Properties</span></span>

| <span data-ttu-id="e7efd-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="e7efd-114">Property</span></span> | <span data-ttu-id="e7efd-115">Способов</span><span class="sxs-lookup"><span data-stu-id="e7efd-115">Modes</span></span> | <span data-ttu-id="e7efd-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e7efd-116">Return type</span></span> | <span data-ttu-id="e7efd-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="e7efd-117">Minimum</span></span><br><span data-ttu-id="e7efd-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="e7efd-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e7efd-119">контекст</span><span class="sxs-lookup"><span data-stu-id="e7efd-119">context</span></span>](office.context.md) | <span data-ttu-id="e7efd-120">Создание</span><span class="sxs-lookup"><span data-stu-id="e7efd-120">Compose</span></span><br><span data-ttu-id="e7efd-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-121">Read</span></span> | [<span data-ttu-id="e7efd-122">Context</span><span class="sxs-lookup"><span data-stu-id="e7efd-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="e7efd-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="e7efd-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="e7efd-124">Enumerations</span></span>

| <span data-ttu-id="e7efd-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="e7efd-125">Enumeration</span></span> | <span data-ttu-id="e7efd-126">Способов</span><span class="sxs-lookup"><span data-stu-id="e7efd-126">Modes</span></span> | <span data-ttu-id="e7efd-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e7efd-127">Return type</span></span> | <span data-ttu-id="e7efd-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="e7efd-128">Minimum</span></span><br><span data-ttu-id="e7efd-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="e7efd-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e7efd-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e7efd-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e7efd-131">Создание</span><span class="sxs-lookup"><span data-stu-id="e7efd-131">Compose</span></span><br><span data-ttu-id="e7efd-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-132">Read</span></span> | <span data-ttu-id="e7efd-133">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-133">String</span></span> | [<span data-ttu-id="e7efd-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e7efd-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e7efd-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e7efd-136">Создание</span><span class="sxs-lookup"><span data-stu-id="e7efd-136">Compose</span></span><br><span data-ttu-id="e7efd-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-137">Read</span></span> | <span data-ttu-id="e7efd-138">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-138">String</span></span> | [<span data-ttu-id="e7efd-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e7efd-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e7efd-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e7efd-141">Создание</span><span class="sxs-lookup"><span data-stu-id="e7efd-141">Compose</span></span><br><span data-ttu-id="e7efd-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-142">Read</span></span> | <span data-ttu-id="e7efd-143">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-143">String</span></span> | [<span data-ttu-id="e7efd-144">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="e7efd-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="e7efd-145">Namespaces</span></span>

<span data-ttu-id="e7efd-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="e7efd-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e7efd-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="e7efd-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e7efd-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="e7efd-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7efd-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e7efd-150">Тип</span><span class="sxs-lookup"><span data-stu-id="e7efd-150">Type</span></span>

*   <span data-ttu-id="e7efd-151">String</span><span class="sxs-lookup"><span data-stu-id="e7efd-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7efd-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7efd-152">Properties:</span></span>

|<span data-ttu-id="e7efd-153">Имя</span><span class="sxs-lookup"><span data-stu-id="e7efd-153">Name</span></span>| <span data-ttu-id="e7efd-154">Тип</span><span class="sxs-lookup"><span data-stu-id="e7efd-154">Type</span></span>| <span data-ttu-id="e7efd-155">Описание</span><span class="sxs-lookup"><span data-stu-id="e7efd-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e7efd-156">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-156">String</span></span>|<span data-ttu-id="e7efd-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="e7efd-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e7efd-158">Для указания</span><span class="sxs-lookup"><span data-stu-id="e7efd-158">String</span></span>|<span data-ttu-id="e7efd-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="e7efd-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7efd-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7efd-160">Requirements</span></span>

|<span data-ttu-id="e7efd-161">Требование</span><span class="sxs-lookup"><span data-stu-id="e7efd-161">Requirement</span></span>| <span data-ttu-id="e7efd-162">Значение</span><span class="sxs-lookup"><span data-stu-id="e7efd-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7efd-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7efd-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e7efd-164">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-164">1.1</span></span>|
|[<span data-ttu-id="e7efd-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7efd-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e7efd-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e7efd-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-167">CoercionType: String</span></span>

<span data-ttu-id="e7efd-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e7efd-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7efd-169">Тип</span><span class="sxs-lookup"><span data-stu-id="e7efd-169">Type</span></span>

*   <span data-ttu-id="e7efd-170">String</span><span class="sxs-lookup"><span data-stu-id="e7efd-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7efd-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7efd-171">Properties:</span></span>

|<span data-ttu-id="e7efd-172">Имя</span><span class="sxs-lookup"><span data-stu-id="e7efd-172">Name</span></span>| <span data-ttu-id="e7efd-173">Тип</span><span class="sxs-lookup"><span data-stu-id="e7efd-173">Type</span></span>| <span data-ttu-id="e7efd-174">Описание</span><span class="sxs-lookup"><span data-stu-id="e7efd-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e7efd-175">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-175">String</span></span>|<span data-ttu-id="e7efd-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="e7efd-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e7efd-177">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-177">String</span></span>|<span data-ttu-id="e7efd-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="e7efd-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7efd-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7efd-179">Requirements</span></span>

|<span data-ttu-id="e7efd-180">Требование</span><span class="sxs-lookup"><span data-stu-id="e7efd-180">Requirement</span></span>| <span data-ttu-id="e7efd-181">Значение</span><span class="sxs-lookup"><span data-stu-id="e7efd-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7efd-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7efd-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e7efd-183">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-183">1.1</span></span>|
|[<span data-ttu-id="e7efd-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7efd-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e7efd-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e7efd-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-186">SourceProperty: String</span></span>

<span data-ttu-id="e7efd-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e7efd-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7efd-188">Тип</span><span class="sxs-lookup"><span data-stu-id="e7efd-188">Type</span></span>

*   <span data-ttu-id="e7efd-189">String</span><span class="sxs-lookup"><span data-stu-id="e7efd-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7efd-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7efd-190">Properties:</span></span>

|<span data-ttu-id="e7efd-191">Имя</span><span class="sxs-lookup"><span data-stu-id="e7efd-191">Name</span></span>| <span data-ttu-id="e7efd-192">Тип</span><span class="sxs-lookup"><span data-stu-id="e7efd-192">Type</span></span>| <span data-ttu-id="e7efd-193">Описание</span><span class="sxs-lookup"><span data-stu-id="e7efd-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e7efd-194">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-194">String</span></span>|<span data-ttu-id="e7efd-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7efd-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e7efd-196">Строка</span><span class="sxs-lookup"><span data-stu-id="e7efd-196">String</span></span>|<span data-ttu-id="e7efd-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7efd-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7efd-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7efd-198">Requirements</span></span>

|<span data-ttu-id="e7efd-199">Требование</span><span class="sxs-lookup"><span data-stu-id="e7efd-199">Requirement</span></span>| <span data-ttu-id="e7efd-200">Значение</span><span class="sxs-lookup"><span data-stu-id="e7efd-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7efd-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7efd-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e7efd-202">1.1</span><span class="sxs-lookup"><span data-stu-id="e7efd-202">1.1</span></span>|
|[<span data-ttu-id="e7efd-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7efd-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e7efd-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7efd-204">Compose or Read</span></span>|
