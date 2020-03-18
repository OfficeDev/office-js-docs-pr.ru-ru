---
title: Пространство имен Office — набор обязательных элементов 1,6
description: Объектная модель для пространства имен верхнего уровня API надстроек Outlook (версия API почтовых ящиков 1,6).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ae2f863e054016636ebffc3ff3925cee018036a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717651"
---
# <a name="office"></a><span data-ttu-id="296af-103">Office</span><span class="sxs-lookup"><span data-stu-id="296af-103">Office</span></span>

<span data-ttu-id="296af-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="296af-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="296af-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="296af-106">Requirements</span></span>

|<span data-ttu-id="296af-107">Требование</span><span class="sxs-lookup"><span data-stu-id="296af-107">Requirement</span></span>| <span data-ttu-id="296af-108">Значение</span><span class="sxs-lookup"><span data-stu-id="296af-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="296af-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="296af-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="296af-110">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-110">1.1</span></span>|
|[<span data-ttu-id="296af-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="296af-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="296af-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="296af-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="296af-113">Properties</span><span class="sxs-lookup"><span data-stu-id="296af-113">Properties</span></span>

| <span data-ttu-id="296af-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="296af-114">Property</span></span> | <span data-ttu-id="296af-115">Способов</span><span class="sxs-lookup"><span data-stu-id="296af-115">Modes</span></span> | <span data-ttu-id="296af-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="296af-116">Return type</span></span> | <span data-ttu-id="296af-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="296af-117">Minimum</span></span><br><span data-ttu-id="296af-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="296af-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="296af-119">контекст</span><span class="sxs-lookup"><span data-stu-id="296af-119">context</span></span>](office.context.md) | <span data-ttu-id="296af-120">Создание</span><span class="sxs-lookup"><span data-stu-id="296af-120">Compose</span></span><br><span data-ttu-id="296af-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="296af-121">Read</span></span> | [<span data-ttu-id="296af-122">Context</span><span class="sxs-lookup"><span data-stu-id="296af-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="296af-123">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="296af-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="296af-124">Enumerations</span></span>

| <span data-ttu-id="296af-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="296af-125">Enumeration</span></span> | <span data-ttu-id="296af-126">Способов</span><span class="sxs-lookup"><span data-stu-id="296af-126">Modes</span></span> | <span data-ttu-id="296af-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="296af-127">Return type</span></span> | <span data-ttu-id="296af-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="296af-128">Minimum</span></span><br><span data-ttu-id="296af-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="296af-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="296af-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="296af-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="296af-131">Создание</span><span class="sxs-lookup"><span data-stu-id="296af-131">Compose</span></span><br><span data-ttu-id="296af-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="296af-132">Read</span></span> | <span data-ttu-id="296af-133">String</span><span class="sxs-lookup"><span data-stu-id="296af-133">String</span></span> | [<span data-ttu-id="296af-134">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="296af-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="296af-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="296af-136">Создание</span><span class="sxs-lookup"><span data-stu-id="296af-136">Compose</span></span><br><span data-ttu-id="296af-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="296af-137">Read</span></span> | <span data-ttu-id="296af-138">String</span><span class="sxs-lookup"><span data-stu-id="296af-138">String</span></span> | [<span data-ttu-id="296af-139">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="296af-140">EventType</span><span class="sxs-lookup"><span data-stu-id="296af-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="296af-141">Создание</span><span class="sxs-lookup"><span data-stu-id="296af-141">Compose</span></span><br><span data-ttu-id="296af-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="296af-142">Read</span></span> | <span data-ttu-id="296af-143">String</span><span class="sxs-lookup"><span data-stu-id="296af-143">String</span></span> | [<span data-ttu-id="296af-144">1,5</span><span class="sxs-lookup"><span data-stu-id="296af-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="296af-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="296af-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="296af-146">Создание</span><span class="sxs-lookup"><span data-stu-id="296af-146">Compose</span></span><br><span data-ttu-id="296af-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="296af-147">Read</span></span> | <span data-ttu-id="296af-148">String</span><span class="sxs-lookup"><span data-stu-id="296af-148">String</span></span> | [<span data-ttu-id="296af-149">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="296af-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="296af-150">Namespaces</span></span>

<span data-ttu-id="296af-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="296af-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="296af-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="296af-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="296af-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="296af-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="296af-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="296af-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="296af-155">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-155">Type</span></span>

*   <span data-ttu-id="296af-156">String</span><span class="sxs-lookup"><span data-stu-id="296af-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="296af-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="296af-157">Properties:</span></span>

|<span data-ttu-id="296af-158">Имя</span><span class="sxs-lookup"><span data-stu-id="296af-158">Name</span></span>| <span data-ttu-id="296af-159">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-159">Type</span></span>| <span data-ttu-id="296af-160">Описание</span><span class="sxs-lookup"><span data-stu-id="296af-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="296af-161">String</span><span class="sxs-lookup"><span data-stu-id="296af-161">String</span></span>|<span data-ttu-id="296af-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="296af-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="296af-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="296af-163">String</span></span>|<span data-ttu-id="296af-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="296af-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="296af-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="296af-165">Requirements</span></span>

|<span data-ttu-id="296af-166">Требование</span><span class="sxs-lookup"><span data-stu-id="296af-166">Requirement</span></span>| <span data-ttu-id="296af-167">Значение</span><span class="sxs-lookup"><span data-stu-id="296af-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="296af-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="296af-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="296af-169">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-169">1.1</span></span>|
|[<span data-ttu-id="296af-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="296af-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="296af-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="296af-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="296af-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="296af-172">CoercionType: String</span></span>

<span data-ttu-id="296af-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="296af-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="296af-174">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-174">Type</span></span>

*   <span data-ttu-id="296af-175">String</span><span class="sxs-lookup"><span data-stu-id="296af-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="296af-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="296af-176">Properties:</span></span>

|<span data-ttu-id="296af-177">Имя</span><span class="sxs-lookup"><span data-stu-id="296af-177">Name</span></span>| <span data-ttu-id="296af-178">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-178">Type</span></span>| <span data-ttu-id="296af-179">Описание</span><span class="sxs-lookup"><span data-stu-id="296af-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="296af-180">String</span><span class="sxs-lookup"><span data-stu-id="296af-180">String</span></span>|<span data-ttu-id="296af-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="296af-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="296af-182">String</span><span class="sxs-lookup"><span data-stu-id="296af-182">String</span></span>|<span data-ttu-id="296af-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="296af-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="296af-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="296af-184">Requirements</span></span>

|<span data-ttu-id="296af-185">Требование</span><span class="sxs-lookup"><span data-stu-id="296af-185">Requirement</span></span>| <span data-ttu-id="296af-186">Значение</span><span class="sxs-lookup"><span data-stu-id="296af-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="296af-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="296af-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="296af-188">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-188">1.1</span></span>|
|[<span data-ttu-id="296af-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="296af-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="296af-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="296af-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="296af-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="296af-191">EventType: String</span></span>

<span data-ttu-id="296af-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="296af-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="296af-193">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-193">Type</span></span>

*   <span data-ttu-id="296af-194">String</span><span class="sxs-lookup"><span data-stu-id="296af-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="296af-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="296af-195">Properties:</span></span>

| <span data-ttu-id="296af-196">Имя</span><span class="sxs-lookup"><span data-stu-id="296af-196">Name</span></span> | <span data-ttu-id="296af-197">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-197">Type</span></span> | <span data-ttu-id="296af-198">Описание</span><span class="sxs-lookup"><span data-stu-id="296af-198">Description</span></span> | <span data-ttu-id="296af-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="296af-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="296af-200">String</span><span class="sxs-lookup"><span data-stu-id="296af-200">String</span></span> | <span data-ttu-id="296af-201">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="296af-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="296af-202">1.5</span><span class="sxs-lookup"><span data-stu-id="296af-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="296af-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="296af-203">Requirements</span></span>

|<span data-ttu-id="296af-204">Требование</span><span class="sxs-lookup"><span data-stu-id="296af-204">Requirement</span></span>| <span data-ttu-id="296af-205">Значение</span><span class="sxs-lookup"><span data-stu-id="296af-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="296af-206">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="296af-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="296af-207">1.5</span><span class="sxs-lookup"><span data-stu-id="296af-207">1.5</span></span> |
|[<span data-ttu-id="296af-208">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="296af-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="296af-209">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="296af-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="296af-210">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="296af-210">SourceProperty: String</span></span>

<span data-ttu-id="296af-211">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="296af-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="296af-212">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-212">Type</span></span>

*   <span data-ttu-id="296af-213">String</span><span class="sxs-lookup"><span data-stu-id="296af-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="296af-214">Свойства:</span><span class="sxs-lookup"><span data-stu-id="296af-214">Properties:</span></span>

|<span data-ttu-id="296af-215">Имя</span><span class="sxs-lookup"><span data-stu-id="296af-215">Name</span></span>| <span data-ttu-id="296af-216">Тип</span><span class="sxs-lookup"><span data-stu-id="296af-216">Type</span></span>| <span data-ttu-id="296af-217">Описание</span><span class="sxs-lookup"><span data-stu-id="296af-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="296af-218">String</span><span class="sxs-lookup"><span data-stu-id="296af-218">String</span></span>|<span data-ttu-id="296af-219">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="296af-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="296af-220">String</span><span class="sxs-lookup"><span data-stu-id="296af-220">String</span></span>|<span data-ttu-id="296af-221">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="296af-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="296af-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="296af-222">Requirements</span></span>

|<span data-ttu-id="296af-223">Требование</span><span class="sxs-lookup"><span data-stu-id="296af-223">Requirement</span></span>| <span data-ttu-id="296af-224">Значение</span><span class="sxs-lookup"><span data-stu-id="296af-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="296af-225">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="296af-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="296af-226">1.1</span><span class="sxs-lookup"><span data-stu-id="296af-226">1.1</span></span>|
|[<span data-ttu-id="296af-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="296af-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="296af-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="296af-228">Compose or Read</span></span>|
