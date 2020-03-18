---
title: Пространство имен Office — набор обязательных элементов 1,2
description: Объектная модель для пространства имен верхнего уровня API надстроек Outlook (версия API почтовых ящиков 1,2).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 10445204d3007d816ebed74ede9eeab5d3dfd83c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720164"
---
# <a name="office"></a><span data-ttu-id="cc212-103">Office</span><span class="sxs-lookup"><span data-stu-id="cc212-103">Office</span></span>

<span data-ttu-id="cc212-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="cc212-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc212-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc212-106">Requirements</span></span>

|<span data-ttu-id="cc212-107">Требование</span><span class="sxs-lookup"><span data-stu-id="cc212-107">Requirement</span></span>| <span data-ttu-id="cc212-108">Значение</span><span class="sxs-lookup"><span data-stu-id="cc212-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc212-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cc212-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cc212-110">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-110">1.1</span></span>|
|[<span data-ttu-id="cc212-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cc212-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cc212-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cc212-113">Properties</span><span class="sxs-lookup"><span data-stu-id="cc212-113">Properties</span></span>

| <span data-ttu-id="cc212-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="cc212-114">Property</span></span> | <span data-ttu-id="cc212-115">Способов</span><span class="sxs-lookup"><span data-stu-id="cc212-115">Modes</span></span> | <span data-ttu-id="cc212-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="cc212-116">Return type</span></span> | <span data-ttu-id="cc212-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="cc212-117">Minimum</span></span><br><span data-ttu-id="cc212-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="cc212-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cc212-119">контекст</span><span class="sxs-lookup"><span data-stu-id="cc212-119">context</span></span>](office.context.md) | <span data-ttu-id="cc212-120">Создание</span><span class="sxs-lookup"><span data-stu-id="cc212-120">Compose</span></span><br><span data-ttu-id="cc212-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-121">Read</span></span> | [<span data-ttu-id="cc212-122">Context</span><span class="sxs-lookup"><span data-stu-id="cc212-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="cc212-123">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cc212-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="cc212-124">Enumerations</span></span>

| <span data-ttu-id="cc212-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="cc212-125">Enumeration</span></span> | <span data-ttu-id="cc212-126">Способов</span><span class="sxs-lookup"><span data-stu-id="cc212-126">Modes</span></span> | <span data-ttu-id="cc212-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="cc212-127">Return type</span></span> | <span data-ttu-id="cc212-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="cc212-128">Minimum</span></span><br><span data-ttu-id="cc212-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="cc212-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cc212-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cc212-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cc212-131">Создание</span><span class="sxs-lookup"><span data-stu-id="cc212-131">Compose</span></span><br><span data-ttu-id="cc212-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-132">Read</span></span> | <span data-ttu-id="cc212-133">String</span><span class="sxs-lookup"><span data-stu-id="cc212-133">String</span></span> | [<span data-ttu-id="cc212-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cc212-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cc212-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cc212-136">Создание</span><span class="sxs-lookup"><span data-stu-id="cc212-136">Compose</span></span><br><span data-ttu-id="cc212-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-137">Read</span></span> | <span data-ttu-id="cc212-138">String</span><span class="sxs-lookup"><span data-stu-id="cc212-138">String</span></span> | [<span data-ttu-id="cc212-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cc212-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cc212-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cc212-141">Создание</span><span class="sxs-lookup"><span data-stu-id="cc212-141">Compose</span></span><br><span data-ttu-id="cc212-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-142">Read</span></span> | <span data-ttu-id="cc212-143">String</span><span class="sxs-lookup"><span data-stu-id="cc212-143">String</span></span> | [<span data-ttu-id="cc212-144">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cc212-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="cc212-145">Namespaces</span></span>

<span data-ttu-id="cc212-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="cc212-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cc212-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="cc212-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cc212-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="cc212-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="cc212-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="cc212-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cc212-150">Тип</span><span class="sxs-lookup"><span data-stu-id="cc212-150">Type</span></span>

*   <span data-ttu-id="cc212-151">String</span><span class="sxs-lookup"><span data-stu-id="cc212-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cc212-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cc212-152">Properties:</span></span>

|<span data-ttu-id="cc212-153">Имя</span><span class="sxs-lookup"><span data-stu-id="cc212-153">Name</span></span>| <span data-ttu-id="cc212-154">Тип</span><span class="sxs-lookup"><span data-stu-id="cc212-154">Type</span></span>| <span data-ttu-id="cc212-155">Описание</span><span class="sxs-lookup"><span data-stu-id="cc212-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cc212-156">String</span><span class="sxs-lookup"><span data-stu-id="cc212-156">String</span></span>|<span data-ttu-id="cc212-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="cc212-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cc212-158">Для указания</span><span class="sxs-lookup"><span data-stu-id="cc212-158">String</span></span>|<span data-ttu-id="cc212-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="cc212-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc212-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc212-160">Requirements</span></span>

|<span data-ttu-id="cc212-161">Требование</span><span class="sxs-lookup"><span data-stu-id="cc212-161">Requirement</span></span>| <span data-ttu-id="cc212-162">Значение</span><span class="sxs-lookup"><span data-stu-id="cc212-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc212-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cc212-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cc212-164">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-164">1.1</span></span>|
|[<span data-ttu-id="cc212-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cc212-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cc212-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cc212-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="cc212-167">CoercionType: String</span></span>

<span data-ttu-id="cc212-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="cc212-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cc212-169">Тип</span><span class="sxs-lookup"><span data-stu-id="cc212-169">Type</span></span>

*   <span data-ttu-id="cc212-170">String</span><span class="sxs-lookup"><span data-stu-id="cc212-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cc212-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cc212-171">Properties:</span></span>

|<span data-ttu-id="cc212-172">Имя</span><span class="sxs-lookup"><span data-stu-id="cc212-172">Name</span></span>| <span data-ttu-id="cc212-173">Тип</span><span class="sxs-lookup"><span data-stu-id="cc212-173">Type</span></span>| <span data-ttu-id="cc212-174">Описание</span><span class="sxs-lookup"><span data-stu-id="cc212-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cc212-175">String</span><span class="sxs-lookup"><span data-stu-id="cc212-175">String</span></span>|<span data-ttu-id="cc212-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="cc212-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cc212-177">String</span><span class="sxs-lookup"><span data-stu-id="cc212-177">String</span></span>|<span data-ttu-id="cc212-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="cc212-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc212-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc212-179">Requirements</span></span>

|<span data-ttu-id="cc212-180">Требование</span><span class="sxs-lookup"><span data-stu-id="cc212-180">Requirement</span></span>| <span data-ttu-id="cc212-181">Значение</span><span class="sxs-lookup"><span data-stu-id="cc212-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc212-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cc212-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cc212-183">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-183">1.1</span></span>|
|[<span data-ttu-id="cc212-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cc212-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cc212-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cc212-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="cc212-186">SourceProperty: String</span></span>

<span data-ttu-id="cc212-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="cc212-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cc212-188">Тип</span><span class="sxs-lookup"><span data-stu-id="cc212-188">Type</span></span>

*   <span data-ttu-id="cc212-189">String</span><span class="sxs-lookup"><span data-stu-id="cc212-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cc212-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cc212-190">Properties:</span></span>

|<span data-ttu-id="cc212-191">Имя</span><span class="sxs-lookup"><span data-stu-id="cc212-191">Name</span></span>| <span data-ttu-id="cc212-192">Тип</span><span class="sxs-lookup"><span data-stu-id="cc212-192">Type</span></span>| <span data-ttu-id="cc212-193">Описание</span><span class="sxs-lookup"><span data-stu-id="cc212-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cc212-194">String</span><span class="sxs-lookup"><span data-stu-id="cc212-194">String</span></span>|<span data-ttu-id="cc212-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="cc212-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cc212-196">String</span><span class="sxs-lookup"><span data-stu-id="cc212-196">String</span></span>|<span data-ttu-id="cc212-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="cc212-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc212-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc212-198">Requirements</span></span>

|<span data-ttu-id="cc212-199">Требование</span><span class="sxs-lookup"><span data-stu-id="cc212-199">Requirement</span></span>| <span data-ttu-id="cc212-200">Значение</span><span class="sxs-lookup"><span data-stu-id="cc212-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc212-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cc212-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cc212-202">1.1</span><span class="sxs-lookup"><span data-stu-id="cc212-202">1.1</span></span>|
|[<span data-ttu-id="cc212-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cc212-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cc212-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cc212-204">Compose or Read</span></span>|
