---
title: Пространство имен Office — набор обязательных элементов 1,1
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f127ab6594d5838700bbc04661d995b01da4f067
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611479"
---
# <a name="office-mailbox-requirement-set-11"></a><span data-ttu-id="76df3-103">Office (набор требований для почтового ящика 1,1)</span><span class="sxs-lookup"><span data-stu-id="76df3-103">Office (Mailbox requirement set 1.1)</span></span>

<span data-ttu-id="76df3-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="76df3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="76df3-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="76df3-106">Requirements</span></span>

|<span data-ttu-id="76df3-107">Требование</span><span class="sxs-lookup"><span data-stu-id="76df3-107">Requirement</span></span>| <span data-ttu-id="76df3-108">Значение</span><span class="sxs-lookup"><span data-stu-id="76df3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="76df3-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76df3-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76df3-110">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-110">1.1</span></span>|
|[<span data-ttu-id="76df3-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76df3-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76df3-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76df3-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="76df3-113">Properties</span><span class="sxs-lookup"><span data-stu-id="76df3-113">Properties</span></span>

| <span data-ttu-id="76df3-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="76df3-114">Property</span></span> | <span data-ttu-id="76df3-115">Способов</span><span class="sxs-lookup"><span data-stu-id="76df3-115">Modes</span></span> | <span data-ttu-id="76df3-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="76df3-116">Return type</span></span> | <span data-ttu-id="76df3-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="76df3-117">Minimum</span></span><br><span data-ttu-id="76df3-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="76df3-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="76df3-119">контекст</span><span class="sxs-lookup"><span data-stu-id="76df3-119">context</span></span>](office.context.md) | <span data-ttu-id="76df3-120">Создание</span><span class="sxs-lookup"><span data-stu-id="76df3-120">Compose</span></span><br><span data-ttu-id="76df3-121">Read</span><span class="sxs-lookup"><span data-stu-id="76df3-121">Read</span></span> | [<span data-ttu-id="76df3-122">Context</span><span class="sxs-lookup"><span data-stu-id="76df3-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="76df3-123">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="76df3-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="76df3-124">Enumerations</span></span>

| <span data-ttu-id="76df3-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="76df3-125">Enumeration</span></span> | <span data-ttu-id="76df3-126">Способов</span><span class="sxs-lookup"><span data-stu-id="76df3-126">Modes</span></span> | <span data-ttu-id="76df3-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="76df3-127">Return type</span></span> | <span data-ttu-id="76df3-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="76df3-128">Minimum</span></span><br><span data-ttu-id="76df3-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="76df3-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="76df3-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="76df3-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="76df3-131">Создание</span><span class="sxs-lookup"><span data-stu-id="76df3-131">Compose</span></span><br><span data-ttu-id="76df3-132">Read</span><span class="sxs-lookup"><span data-stu-id="76df3-132">Read</span></span> | <span data-ttu-id="76df3-133">String</span><span class="sxs-lookup"><span data-stu-id="76df3-133">String</span></span> | [<span data-ttu-id="76df3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="76df3-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="76df3-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="76df3-136">Создание</span><span class="sxs-lookup"><span data-stu-id="76df3-136">Compose</span></span><br><span data-ttu-id="76df3-137">Read</span><span class="sxs-lookup"><span data-stu-id="76df3-137">Read</span></span> | <span data-ttu-id="76df3-138">String</span><span class="sxs-lookup"><span data-stu-id="76df3-138">String</span></span> | [<span data-ttu-id="76df3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="76df3-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="76df3-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="76df3-141">Создание</span><span class="sxs-lookup"><span data-stu-id="76df3-141">Compose</span></span><br><span data-ttu-id="76df3-142">Read</span><span class="sxs-lookup"><span data-stu-id="76df3-142">Read</span></span> | <span data-ttu-id="76df3-143">String</span><span class="sxs-lookup"><span data-stu-id="76df3-143">String</span></span> | [<span data-ttu-id="76df3-144">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="76df3-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="76df3-145">Namespaces</span></span>

<span data-ttu-id="76df3-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="76df3-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="76df3-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="76df3-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="76df3-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="76df3-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="76df3-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="76df3-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="76df3-150">Тип</span><span class="sxs-lookup"><span data-stu-id="76df3-150">Type</span></span>

*   <span data-ttu-id="76df3-151">String</span><span class="sxs-lookup"><span data-stu-id="76df3-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76df3-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76df3-152">Properties:</span></span>

|<span data-ttu-id="76df3-153">Имя</span><span class="sxs-lookup"><span data-stu-id="76df3-153">Name</span></span>| <span data-ttu-id="76df3-154">Тип</span><span class="sxs-lookup"><span data-stu-id="76df3-154">Type</span></span>| <span data-ttu-id="76df3-155">Описание</span><span class="sxs-lookup"><span data-stu-id="76df3-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="76df3-156">String</span><span class="sxs-lookup"><span data-stu-id="76df3-156">String</span></span>|<span data-ttu-id="76df3-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="76df3-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="76df3-158">Для указания</span><span class="sxs-lookup"><span data-stu-id="76df3-158">String</span></span>|<span data-ttu-id="76df3-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="76df3-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76df3-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="76df3-160">Requirements</span></span>

|<span data-ttu-id="76df3-161">Требование</span><span class="sxs-lookup"><span data-stu-id="76df3-161">Requirement</span></span>| <span data-ttu-id="76df3-162">Значение</span><span class="sxs-lookup"><span data-stu-id="76df3-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="76df3-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76df3-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76df3-164">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-164">1.1</span></span>|
|[<span data-ttu-id="76df3-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76df3-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76df3-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76df3-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="76df3-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="76df3-167">CoercionType: String</span></span>

<span data-ttu-id="76df3-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="76df3-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="76df3-169">Тип</span><span class="sxs-lookup"><span data-stu-id="76df3-169">Type</span></span>

*   <span data-ttu-id="76df3-170">String</span><span class="sxs-lookup"><span data-stu-id="76df3-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76df3-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76df3-171">Properties:</span></span>

|<span data-ttu-id="76df3-172">Имя</span><span class="sxs-lookup"><span data-stu-id="76df3-172">Name</span></span>| <span data-ttu-id="76df3-173">Тип</span><span class="sxs-lookup"><span data-stu-id="76df3-173">Type</span></span>| <span data-ttu-id="76df3-174">Описание</span><span class="sxs-lookup"><span data-stu-id="76df3-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="76df3-175">String</span><span class="sxs-lookup"><span data-stu-id="76df3-175">String</span></span>|<span data-ttu-id="76df3-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="76df3-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="76df3-177">String</span><span class="sxs-lookup"><span data-stu-id="76df3-177">String</span></span>|<span data-ttu-id="76df3-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="76df3-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76df3-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="76df3-179">Requirements</span></span>

|<span data-ttu-id="76df3-180">Требование</span><span class="sxs-lookup"><span data-stu-id="76df3-180">Requirement</span></span>| <span data-ttu-id="76df3-181">Значение</span><span class="sxs-lookup"><span data-stu-id="76df3-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="76df3-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76df3-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76df3-183">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-183">1.1</span></span>|
|[<span data-ttu-id="76df3-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76df3-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76df3-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76df3-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="76df3-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="76df3-186">SourceProperty: String</span></span>

<span data-ttu-id="76df3-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="76df3-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="76df3-188">Тип</span><span class="sxs-lookup"><span data-stu-id="76df3-188">Type</span></span>

*   <span data-ttu-id="76df3-189">String</span><span class="sxs-lookup"><span data-stu-id="76df3-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76df3-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76df3-190">Properties:</span></span>

|<span data-ttu-id="76df3-191">Имя</span><span class="sxs-lookup"><span data-stu-id="76df3-191">Name</span></span>| <span data-ttu-id="76df3-192">Тип</span><span class="sxs-lookup"><span data-stu-id="76df3-192">Type</span></span>| <span data-ttu-id="76df3-193">Описание</span><span class="sxs-lookup"><span data-stu-id="76df3-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="76df3-194">String</span><span class="sxs-lookup"><span data-stu-id="76df3-194">String</span></span>|<span data-ttu-id="76df3-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="76df3-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="76df3-196">String</span><span class="sxs-lookup"><span data-stu-id="76df3-196">String</span></span>|<span data-ttu-id="76df3-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="76df3-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76df3-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="76df3-198">Requirements</span></span>

|<span data-ttu-id="76df3-199">Требование</span><span class="sxs-lookup"><span data-stu-id="76df3-199">Requirement</span></span>| <span data-ttu-id="76df3-200">Значение</span><span class="sxs-lookup"><span data-stu-id="76df3-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="76df3-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76df3-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76df3-202">1.1</span><span class="sxs-lookup"><span data-stu-id="76df3-202">1.1</span></span>|
|[<span data-ttu-id="76df3-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76df3-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76df3-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76df3-204">Compose or Read</span></span>|