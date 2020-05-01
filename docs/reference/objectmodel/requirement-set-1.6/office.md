---
title: Пространство имен Office — набор обязательных элементов 1,6
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: dc7f62cc3f01e56f6c05b6cf40a4b73e87aea5e4
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891315"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="aaf27-103">Office (набор требований для почтового ящика 1,6)</span><span class="sxs-lookup"><span data-stu-id="aaf27-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="aaf27-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="aaf27-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="aaf27-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="aaf27-106">Requirements</span></span>

|<span data-ttu-id="aaf27-107">Требование</span><span class="sxs-lookup"><span data-stu-id="aaf27-107">Requirement</span></span>| <span data-ttu-id="aaf27-108">Значение</span><span class="sxs-lookup"><span data-stu-id="aaf27-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="aaf27-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aaf27-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aaf27-110">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-110">1.1</span></span>|
|[<span data-ttu-id="aaf27-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aaf27-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aaf27-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="aaf27-113">Properties</span><span class="sxs-lookup"><span data-stu-id="aaf27-113">Properties</span></span>

| <span data-ttu-id="aaf27-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="aaf27-114">Property</span></span> | <span data-ttu-id="aaf27-115">Способов</span><span class="sxs-lookup"><span data-stu-id="aaf27-115">Modes</span></span> | <span data-ttu-id="aaf27-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="aaf27-116">Return type</span></span> | <span data-ttu-id="aaf27-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="aaf27-117">Minimum</span></span><br><span data-ttu-id="aaf27-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="aaf27-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="aaf27-119">контекст</span><span class="sxs-lookup"><span data-stu-id="aaf27-119">context</span></span>](office.context.md) | <span data-ttu-id="aaf27-120">Создание</span><span class="sxs-lookup"><span data-stu-id="aaf27-120">Compose</span></span><br><span data-ttu-id="aaf27-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-121">Read</span></span> | [<span data-ttu-id="aaf27-122">Context</span><span class="sxs-lookup"><span data-stu-id="aaf27-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="aaf27-123">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="aaf27-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="aaf27-124">Enumerations</span></span>

| <span data-ttu-id="aaf27-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="aaf27-125">Enumeration</span></span> | <span data-ttu-id="aaf27-126">Способов</span><span class="sxs-lookup"><span data-stu-id="aaf27-126">Modes</span></span> | <span data-ttu-id="aaf27-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="aaf27-127">Return type</span></span> | <span data-ttu-id="aaf27-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="aaf27-128">Minimum</span></span><br><span data-ttu-id="aaf27-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="aaf27-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="aaf27-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="aaf27-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="aaf27-131">Создание</span><span class="sxs-lookup"><span data-stu-id="aaf27-131">Compose</span></span><br><span data-ttu-id="aaf27-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-132">Read</span></span> | <span data-ttu-id="aaf27-133">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-133">String</span></span> | [<span data-ttu-id="aaf27-134">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="aaf27-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="aaf27-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="aaf27-136">Создание</span><span class="sxs-lookup"><span data-stu-id="aaf27-136">Compose</span></span><br><span data-ttu-id="aaf27-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-137">Read</span></span> | <span data-ttu-id="aaf27-138">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-138">String</span></span> | [<span data-ttu-id="aaf27-139">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="aaf27-140">EventType</span><span class="sxs-lookup"><span data-stu-id="aaf27-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="aaf27-141">Создание</span><span class="sxs-lookup"><span data-stu-id="aaf27-141">Compose</span></span><br><span data-ttu-id="aaf27-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-142">Read</span></span> | <span data-ttu-id="aaf27-143">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-143">String</span></span> | [<span data-ttu-id="aaf27-144">1,5</span><span class="sxs-lookup"><span data-stu-id="aaf27-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="aaf27-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="aaf27-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="aaf27-146">Создание</span><span class="sxs-lookup"><span data-stu-id="aaf27-146">Compose</span></span><br><span data-ttu-id="aaf27-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-147">Read</span></span> | <span data-ttu-id="aaf27-148">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-148">String</span></span> | [<span data-ttu-id="aaf27-149">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="aaf27-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="aaf27-150">Namespaces</span></span>

<span data-ttu-id="aaf27-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="aaf27-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="aaf27-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="aaf27-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="aaf27-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="aaf27-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="aaf27-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="aaf27-155">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-155">Type</span></span>

*   <span data-ttu-id="aaf27-156">String</span><span class="sxs-lookup"><span data-stu-id="aaf27-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aaf27-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="aaf27-157">Properties:</span></span>

|<span data-ttu-id="aaf27-158">Имя</span><span class="sxs-lookup"><span data-stu-id="aaf27-158">Name</span></span>| <span data-ttu-id="aaf27-159">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-159">Type</span></span>| <span data-ttu-id="aaf27-160">Описание</span><span class="sxs-lookup"><span data-stu-id="aaf27-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="aaf27-161">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-161">String</span></span>|<span data-ttu-id="aaf27-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="aaf27-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="aaf27-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="aaf27-163">String</span></span>|<span data-ttu-id="aaf27-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="aaf27-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aaf27-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="aaf27-165">Requirements</span></span>

|<span data-ttu-id="aaf27-166">Требование</span><span class="sxs-lookup"><span data-stu-id="aaf27-166">Requirement</span></span>| <span data-ttu-id="aaf27-167">Значение</span><span class="sxs-lookup"><span data-stu-id="aaf27-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="aaf27-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aaf27-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aaf27-169">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-169">1.1</span></span>|
|[<span data-ttu-id="aaf27-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aaf27-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aaf27-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="aaf27-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-172">CoercionType: String</span></span>

<span data-ttu-id="aaf27-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="aaf27-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="aaf27-174">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-174">Type</span></span>

*   <span data-ttu-id="aaf27-175">String</span><span class="sxs-lookup"><span data-stu-id="aaf27-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aaf27-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="aaf27-176">Properties:</span></span>

|<span data-ttu-id="aaf27-177">Имя</span><span class="sxs-lookup"><span data-stu-id="aaf27-177">Name</span></span>| <span data-ttu-id="aaf27-178">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-178">Type</span></span>| <span data-ttu-id="aaf27-179">Описание</span><span class="sxs-lookup"><span data-stu-id="aaf27-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="aaf27-180">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-180">String</span></span>|<span data-ttu-id="aaf27-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="aaf27-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="aaf27-182">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-182">String</span></span>|<span data-ttu-id="aaf27-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="aaf27-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aaf27-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="aaf27-184">Requirements</span></span>

|<span data-ttu-id="aaf27-185">Требование</span><span class="sxs-lookup"><span data-stu-id="aaf27-185">Requirement</span></span>| <span data-ttu-id="aaf27-186">Значение</span><span class="sxs-lookup"><span data-stu-id="aaf27-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="aaf27-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aaf27-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aaf27-188">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-188">1.1</span></span>|
|[<span data-ttu-id="aaf27-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aaf27-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aaf27-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="aaf27-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-191">EventType: String</span></span>

<span data-ttu-id="aaf27-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="aaf27-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="aaf27-193">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-193">Type</span></span>

*   <span data-ttu-id="aaf27-194">String</span><span class="sxs-lookup"><span data-stu-id="aaf27-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aaf27-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="aaf27-195">Properties:</span></span>

| <span data-ttu-id="aaf27-196">Имя</span><span class="sxs-lookup"><span data-stu-id="aaf27-196">Name</span></span> | <span data-ttu-id="aaf27-197">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-197">Type</span></span> | <span data-ttu-id="aaf27-198">Описание</span><span class="sxs-lookup"><span data-stu-id="aaf27-198">Description</span></span> | <span data-ttu-id="aaf27-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="aaf27-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="aaf27-200">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-200">String</span></span> | <span data-ttu-id="aaf27-201">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="aaf27-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="aaf27-202">1.5</span><span class="sxs-lookup"><span data-stu-id="aaf27-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aaf27-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="aaf27-203">Requirements</span></span>

|<span data-ttu-id="aaf27-204">Требование</span><span class="sxs-lookup"><span data-stu-id="aaf27-204">Requirement</span></span>| <span data-ttu-id="aaf27-205">Значение</span><span class="sxs-lookup"><span data-stu-id="aaf27-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="aaf27-206">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aaf27-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aaf27-207">1.5</span><span class="sxs-lookup"><span data-stu-id="aaf27-207">1.5</span></span> |
|[<span data-ttu-id="aaf27-208">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aaf27-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aaf27-209">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="aaf27-210">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-210">SourceProperty: String</span></span>

<span data-ttu-id="aaf27-211">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="aaf27-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="aaf27-212">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-212">Type</span></span>

*   <span data-ttu-id="aaf27-213">String</span><span class="sxs-lookup"><span data-stu-id="aaf27-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aaf27-214">Свойства:</span><span class="sxs-lookup"><span data-stu-id="aaf27-214">Properties:</span></span>

|<span data-ttu-id="aaf27-215">Имя</span><span class="sxs-lookup"><span data-stu-id="aaf27-215">Name</span></span>| <span data-ttu-id="aaf27-216">Тип</span><span class="sxs-lookup"><span data-stu-id="aaf27-216">Type</span></span>| <span data-ttu-id="aaf27-217">Описание</span><span class="sxs-lookup"><span data-stu-id="aaf27-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="aaf27-218">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-218">String</span></span>|<span data-ttu-id="aaf27-219">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="aaf27-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="aaf27-220">Строка</span><span class="sxs-lookup"><span data-stu-id="aaf27-220">String</span></span>|<span data-ttu-id="aaf27-221">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="aaf27-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aaf27-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="aaf27-222">Requirements</span></span>

|<span data-ttu-id="aaf27-223">Требование</span><span class="sxs-lookup"><span data-stu-id="aaf27-223">Requirement</span></span>| <span data-ttu-id="aaf27-224">Значение</span><span class="sxs-lookup"><span data-stu-id="aaf27-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="aaf27-225">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aaf27-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aaf27-226">1.1</span><span class="sxs-lookup"><span data-stu-id="aaf27-226">1.1</span></span>|
|[<span data-ttu-id="aaf27-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aaf27-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aaf27-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aaf27-228">Compose or Read</span></span>|
