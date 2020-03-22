---
title: Пространство имен Office — набор обязательных элементов 1,5
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 848aa30c07b936c8454b2833d5dce3e1d15ee193
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891350"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="9da4a-103">Office (набор требований для почтового ящика 1,5)</span><span class="sxs-lookup"><span data-stu-id="9da4a-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="9da4a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9da4a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9da4a-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="9da4a-106">Requirements</span></span>

|<span data-ttu-id="9da4a-107">Требование</span><span class="sxs-lookup"><span data-stu-id="9da4a-107">Requirement</span></span>| <span data-ttu-id="9da4a-108">Значение</span><span class="sxs-lookup"><span data-stu-id="9da4a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da4a-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9da4a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9da4a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-110">1.1</span></span>|
|[<span data-ttu-id="9da4a-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da4a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9da4a-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="9da4a-113">Properties</span><span class="sxs-lookup"><span data-stu-id="9da4a-113">Properties</span></span>

| <span data-ttu-id="9da4a-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="9da4a-114">Property</span></span> | <span data-ttu-id="9da4a-115">Способов</span><span class="sxs-lookup"><span data-stu-id="9da4a-115">Modes</span></span> | <span data-ttu-id="9da4a-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="9da4a-116">Return type</span></span> | <span data-ttu-id="9da4a-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="9da4a-117">Minimum</span></span><br><span data-ttu-id="9da4a-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="9da4a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9da4a-119">контекст</span><span class="sxs-lookup"><span data-stu-id="9da4a-119">context</span></span>](office.context.md) | <span data-ttu-id="9da4a-120">Создание</span><span class="sxs-lookup"><span data-stu-id="9da4a-120">Compose</span></span><br><span data-ttu-id="9da4a-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-121">Read</span></span> | [<span data-ttu-id="9da4a-122">Context</span><span class="sxs-lookup"><span data-stu-id="9da4a-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="9da4a-123">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="9da4a-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="9da4a-124">Enumerations</span></span>

| <span data-ttu-id="9da4a-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="9da4a-125">Enumeration</span></span> | <span data-ttu-id="9da4a-126">Способов</span><span class="sxs-lookup"><span data-stu-id="9da4a-126">Modes</span></span> | <span data-ttu-id="9da4a-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="9da4a-127">Return type</span></span> | <span data-ttu-id="9da4a-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="9da4a-128">Minimum</span></span><br><span data-ttu-id="9da4a-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="9da4a-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9da4a-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9da4a-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9da4a-131">Создание</span><span class="sxs-lookup"><span data-stu-id="9da4a-131">Compose</span></span><br><span data-ttu-id="9da4a-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-132">Read</span></span> | <span data-ttu-id="9da4a-133">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-133">String</span></span> | [<span data-ttu-id="9da4a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9da4a-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9da4a-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9da4a-136">Создание</span><span class="sxs-lookup"><span data-stu-id="9da4a-136">Compose</span></span><br><span data-ttu-id="9da4a-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-137">Read</span></span> | <span data-ttu-id="9da4a-138">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-138">String</span></span> | [<span data-ttu-id="9da4a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9da4a-140">EventType</span><span class="sxs-lookup"><span data-stu-id="9da4a-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9da4a-141">Создание</span><span class="sxs-lookup"><span data-stu-id="9da4a-141">Compose</span></span><br><span data-ttu-id="9da4a-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-142">Read</span></span> | <span data-ttu-id="9da4a-143">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-143">String</span></span> | [<span data-ttu-id="9da4a-144">1,5</span><span class="sxs-lookup"><span data-stu-id="9da4a-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9da4a-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9da4a-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9da4a-146">Создание</span><span class="sxs-lookup"><span data-stu-id="9da4a-146">Compose</span></span><br><span data-ttu-id="9da4a-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-147">Read</span></span> | <span data-ttu-id="9da4a-148">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-148">String</span></span> | [<span data-ttu-id="9da4a-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="9da4a-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="9da4a-150">Namespaces</span></span>

<span data-ttu-id="9da4a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="9da4a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="9da4a-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="9da4a-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="9da4a-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="9da4a-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="9da4a-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="9da4a-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9da4a-155">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-155">Type</span></span>

*   <span data-ttu-id="9da4a-156">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da4a-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da4a-157">Properties:</span></span>

|<span data-ttu-id="9da4a-158">Имя</span><span class="sxs-lookup"><span data-stu-id="9da4a-158">Name</span></span>| <span data-ttu-id="9da4a-159">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-159">Type</span></span>| <span data-ttu-id="9da4a-160">Описание</span><span class="sxs-lookup"><span data-stu-id="9da4a-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9da4a-161">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-161">String</span></span>|<span data-ttu-id="9da4a-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="9da4a-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9da4a-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="9da4a-163">String</span></span>|<span data-ttu-id="9da4a-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="9da4a-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9da4a-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="9da4a-165">Requirements</span></span>

|<span data-ttu-id="9da4a-166">Требование</span><span class="sxs-lookup"><span data-stu-id="9da4a-166">Requirement</span></span>| <span data-ttu-id="9da4a-167">Значение</span><span class="sxs-lookup"><span data-stu-id="9da4a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da4a-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9da4a-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9da4a-169">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-169">1.1</span></span>|
|[<span data-ttu-id="9da4a-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da4a-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9da4a-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="9da4a-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="9da4a-172">CoercionType: String</span></span>

<span data-ttu-id="9da4a-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9da4a-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9da4a-174">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-174">Type</span></span>

*   <span data-ttu-id="9da4a-175">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da4a-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da4a-176">Properties:</span></span>

|<span data-ttu-id="9da4a-177">Имя</span><span class="sxs-lookup"><span data-stu-id="9da4a-177">Name</span></span>| <span data-ttu-id="9da4a-178">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-178">Type</span></span>| <span data-ttu-id="9da4a-179">Описание</span><span class="sxs-lookup"><span data-stu-id="9da4a-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9da4a-180">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-180">String</span></span>|<span data-ttu-id="9da4a-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="9da4a-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9da4a-182">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-182">String</span></span>|<span data-ttu-id="9da4a-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="9da4a-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9da4a-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="9da4a-184">Requirements</span></span>

|<span data-ttu-id="9da4a-185">Требование</span><span class="sxs-lookup"><span data-stu-id="9da4a-185">Requirement</span></span>| <span data-ttu-id="9da4a-186">Значение</span><span class="sxs-lookup"><span data-stu-id="9da4a-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da4a-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9da4a-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9da4a-188">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-188">1.1</span></span>|
|[<span data-ttu-id="9da4a-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da4a-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9da4a-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="9da4a-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="9da4a-191">EventType: String</span></span>

<span data-ttu-id="9da4a-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="9da4a-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9da4a-193">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-193">Type</span></span>

*   <span data-ttu-id="9da4a-194">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da4a-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da4a-195">Properties:</span></span>

| <span data-ttu-id="9da4a-196">Имя</span><span class="sxs-lookup"><span data-stu-id="9da4a-196">Name</span></span> | <span data-ttu-id="9da4a-197">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-197">Type</span></span> | <span data-ttu-id="9da4a-198">Описание</span><span class="sxs-lookup"><span data-stu-id="9da4a-198">Description</span></span> | <span data-ttu-id="9da4a-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="9da4a-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="9da4a-200">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-200">String</span></span> | <span data-ttu-id="9da4a-201">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="9da4a-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="9da4a-202">1.5</span><span class="sxs-lookup"><span data-stu-id="9da4a-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9da4a-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="9da4a-203">Requirements</span></span>

|<span data-ttu-id="9da4a-204">Требование</span><span class="sxs-lookup"><span data-stu-id="9da4a-204">Requirement</span></span>| <span data-ttu-id="9da4a-205">Значение</span><span class="sxs-lookup"><span data-stu-id="9da4a-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da4a-206">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9da4a-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9da4a-207">1.5</span><span class="sxs-lookup"><span data-stu-id="9da4a-207">1.5</span></span> |
|[<span data-ttu-id="9da4a-208">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da4a-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9da4a-209">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="9da4a-210">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="9da4a-210">SourceProperty: String</span></span>

<span data-ttu-id="9da4a-211">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9da4a-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9da4a-212">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-212">Type</span></span>

*   <span data-ttu-id="9da4a-213">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da4a-214">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da4a-214">Properties:</span></span>

|<span data-ttu-id="9da4a-215">Имя</span><span class="sxs-lookup"><span data-stu-id="9da4a-215">Name</span></span>| <span data-ttu-id="9da4a-216">Тип</span><span class="sxs-lookup"><span data-stu-id="9da4a-216">Type</span></span>| <span data-ttu-id="9da4a-217">Описание</span><span class="sxs-lookup"><span data-stu-id="9da4a-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9da4a-218">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-218">String</span></span>|<span data-ttu-id="9da4a-219">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="9da4a-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9da4a-220">String</span><span class="sxs-lookup"><span data-stu-id="9da4a-220">String</span></span>|<span data-ttu-id="9da4a-221">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="9da4a-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9da4a-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="9da4a-222">Requirements</span></span>

|<span data-ttu-id="9da4a-223">Требование</span><span class="sxs-lookup"><span data-stu-id="9da4a-223">Requirement</span></span>| <span data-ttu-id="9da4a-224">Значение</span><span class="sxs-lookup"><span data-stu-id="9da4a-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da4a-225">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9da4a-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9da4a-226">1.1</span><span class="sxs-lookup"><span data-stu-id="9da4a-226">1.1</span></span>|
|[<span data-ttu-id="9da4a-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da4a-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9da4a-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9da4a-228">Compose or Read</span></span>|
