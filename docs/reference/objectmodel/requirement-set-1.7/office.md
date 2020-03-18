---
title: Пространство имен Office — набор обязательных элементов 1,7
description: Это пространство имен предоставляет общие интерфейсы, используемые надстройками Outlook для Office (набор требований 1,7).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 50fa22ac14aee3b7276be83813db248681435dc1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717601"
---
# <a name="office"></a><span data-ttu-id="32144-103">Office</span><span class="sxs-lookup"><span data-stu-id="32144-103">Office</span></span>

<span data-ttu-id="32144-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="32144-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="32144-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="32144-106">Requirements</span></span>

|<span data-ttu-id="32144-107">Требование</span><span class="sxs-lookup"><span data-stu-id="32144-107">Requirement</span></span>| <span data-ttu-id="32144-108">Значение</span><span class="sxs-lookup"><span data-stu-id="32144-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="32144-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32144-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32144-110">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-110">1.1</span></span>|
|[<span data-ttu-id="32144-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32144-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32144-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32144-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="32144-113">Properties</span><span class="sxs-lookup"><span data-stu-id="32144-113">Properties</span></span>

| <span data-ttu-id="32144-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="32144-114">Property</span></span> | <span data-ttu-id="32144-115">Способов</span><span class="sxs-lookup"><span data-stu-id="32144-115">Modes</span></span> | <span data-ttu-id="32144-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="32144-116">Return type</span></span> | <span data-ttu-id="32144-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="32144-117">Minimum</span></span><br><span data-ttu-id="32144-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="32144-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="32144-119">контекст</span><span class="sxs-lookup"><span data-stu-id="32144-119">context</span></span>](office.context.md) | <span data-ttu-id="32144-120">Создание</span><span class="sxs-lookup"><span data-stu-id="32144-120">Compose</span></span><br><span data-ttu-id="32144-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="32144-121">Read</span></span> | [<span data-ttu-id="32144-122">Context</span><span class="sxs-lookup"><span data-stu-id="32144-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="32144-123">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="32144-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="32144-124">Enumerations</span></span>

| <span data-ttu-id="32144-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="32144-125">Enumeration</span></span> | <span data-ttu-id="32144-126">Способов</span><span class="sxs-lookup"><span data-stu-id="32144-126">Modes</span></span> | <span data-ttu-id="32144-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="32144-127">Return type</span></span> | <span data-ttu-id="32144-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="32144-128">Minimum</span></span><br><span data-ttu-id="32144-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="32144-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="32144-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="32144-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="32144-131">Создание</span><span class="sxs-lookup"><span data-stu-id="32144-131">Compose</span></span><br><span data-ttu-id="32144-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="32144-132">Read</span></span> | <span data-ttu-id="32144-133">String</span><span class="sxs-lookup"><span data-stu-id="32144-133">String</span></span> | [<span data-ttu-id="32144-134">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32144-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="32144-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="32144-136">Создание</span><span class="sxs-lookup"><span data-stu-id="32144-136">Compose</span></span><br><span data-ttu-id="32144-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="32144-137">Read</span></span> | <span data-ttu-id="32144-138">String</span><span class="sxs-lookup"><span data-stu-id="32144-138">String</span></span> | [<span data-ttu-id="32144-139">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32144-140">EventType</span><span class="sxs-lookup"><span data-stu-id="32144-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="32144-141">Создание</span><span class="sxs-lookup"><span data-stu-id="32144-141">Compose</span></span><br><span data-ttu-id="32144-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="32144-142">Read</span></span> | <span data-ttu-id="32144-143">String</span><span class="sxs-lookup"><span data-stu-id="32144-143">String</span></span> | [<span data-ttu-id="32144-144">1,5</span><span class="sxs-lookup"><span data-stu-id="32144-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="32144-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="32144-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="32144-146">Создание</span><span class="sxs-lookup"><span data-stu-id="32144-146">Compose</span></span><br><span data-ttu-id="32144-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="32144-147">Read</span></span> | <span data-ttu-id="32144-148">String</span><span class="sxs-lookup"><span data-stu-id="32144-148">String</span></span> | [<span data-ttu-id="32144-149">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="32144-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="32144-150">Namespaces</span></span>

<span data-ttu-id="32144-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="32144-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="32144-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="32144-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="32144-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="32144-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="32144-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="32144-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="32144-155">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-155">Type</span></span>

*   <span data-ttu-id="32144-156">String</span><span class="sxs-lookup"><span data-stu-id="32144-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32144-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32144-157">Properties:</span></span>

|<span data-ttu-id="32144-158">Имя</span><span class="sxs-lookup"><span data-stu-id="32144-158">Name</span></span>| <span data-ttu-id="32144-159">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-159">Type</span></span>| <span data-ttu-id="32144-160">Описание</span><span class="sxs-lookup"><span data-stu-id="32144-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="32144-161">String</span><span class="sxs-lookup"><span data-stu-id="32144-161">String</span></span>|<span data-ttu-id="32144-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="32144-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="32144-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="32144-163">String</span></span>|<span data-ttu-id="32144-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="32144-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32144-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="32144-165">Requirements</span></span>

|<span data-ttu-id="32144-166">Требование</span><span class="sxs-lookup"><span data-stu-id="32144-166">Requirement</span></span>| <span data-ttu-id="32144-167">Значение</span><span class="sxs-lookup"><span data-stu-id="32144-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="32144-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32144-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32144-169">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-169">1.1</span></span>|
|[<span data-ttu-id="32144-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32144-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32144-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32144-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="32144-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="32144-172">CoercionType: String</span></span>

<span data-ttu-id="32144-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="32144-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="32144-174">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-174">Type</span></span>

*   <span data-ttu-id="32144-175">String</span><span class="sxs-lookup"><span data-stu-id="32144-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32144-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32144-176">Properties:</span></span>

|<span data-ttu-id="32144-177">Имя</span><span class="sxs-lookup"><span data-stu-id="32144-177">Name</span></span>| <span data-ttu-id="32144-178">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-178">Type</span></span>| <span data-ttu-id="32144-179">Описание</span><span class="sxs-lookup"><span data-stu-id="32144-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="32144-180">String</span><span class="sxs-lookup"><span data-stu-id="32144-180">String</span></span>|<span data-ttu-id="32144-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="32144-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="32144-182">String</span><span class="sxs-lookup"><span data-stu-id="32144-182">String</span></span>|<span data-ttu-id="32144-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="32144-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32144-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="32144-184">Requirements</span></span>

|<span data-ttu-id="32144-185">Требование</span><span class="sxs-lookup"><span data-stu-id="32144-185">Requirement</span></span>| <span data-ttu-id="32144-186">Значение</span><span class="sxs-lookup"><span data-stu-id="32144-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="32144-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32144-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32144-188">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-188">1.1</span></span>|
|[<span data-ttu-id="32144-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32144-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32144-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32144-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="32144-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="32144-191">EventType: String</span></span>

<span data-ttu-id="32144-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="32144-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="32144-193">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-193">Type</span></span>

*   <span data-ttu-id="32144-194">String</span><span class="sxs-lookup"><span data-stu-id="32144-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32144-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32144-195">Properties:</span></span>

| <span data-ttu-id="32144-196">Имя</span><span class="sxs-lookup"><span data-stu-id="32144-196">Name</span></span> | <span data-ttu-id="32144-197">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-197">Type</span></span> | <span data-ttu-id="32144-198">Описание</span><span class="sxs-lookup"><span data-stu-id="32144-198">Description</span></span> | <span data-ttu-id="32144-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="32144-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="32144-200">String</span><span class="sxs-lookup"><span data-stu-id="32144-200">String</span></span> | <span data-ttu-id="32144-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="32144-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="32144-202">1.7</span><span class="sxs-lookup"><span data-stu-id="32144-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="32144-203">String</span><span class="sxs-lookup"><span data-stu-id="32144-203">String</span></span> | <span data-ttu-id="32144-204">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="32144-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="32144-205">1.5</span><span class="sxs-lookup"><span data-stu-id="32144-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="32144-206">String</span><span class="sxs-lookup"><span data-stu-id="32144-206">String</span></span> | <span data-ttu-id="32144-207">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="32144-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="32144-208">1.7</span><span class="sxs-lookup"><span data-stu-id="32144-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="32144-209">String</span><span class="sxs-lookup"><span data-stu-id="32144-209">String</span></span> | <span data-ttu-id="32144-210">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="32144-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="32144-211">1.7</span><span class="sxs-lookup"><span data-stu-id="32144-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="32144-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="32144-212">Requirements</span></span>

|<span data-ttu-id="32144-213">Требование</span><span class="sxs-lookup"><span data-stu-id="32144-213">Requirement</span></span>| <span data-ttu-id="32144-214">Значение</span><span class="sxs-lookup"><span data-stu-id="32144-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="32144-215">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="32144-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32144-216">1.5</span><span class="sxs-lookup"><span data-stu-id="32144-216">1.5</span></span> |
|[<span data-ttu-id="32144-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32144-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32144-218">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32144-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="32144-219">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="32144-219">SourceProperty: String</span></span>

<span data-ttu-id="32144-220">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="32144-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="32144-221">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-221">Type</span></span>

*   <span data-ttu-id="32144-222">String</span><span class="sxs-lookup"><span data-stu-id="32144-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32144-223">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32144-223">Properties:</span></span>

|<span data-ttu-id="32144-224">Имя</span><span class="sxs-lookup"><span data-stu-id="32144-224">Name</span></span>| <span data-ttu-id="32144-225">Тип</span><span class="sxs-lookup"><span data-stu-id="32144-225">Type</span></span>| <span data-ttu-id="32144-226">Описание</span><span class="sxs-lookup"><span data-stu-id="32144-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="32144-227">String</span><span class="sxs-lookup"><span data-stu-id="32144-227">String</span></span>|<span data-ttu-id="32144-228">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="32144-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="32144-229">String</span><span class="sxs-lookup"><span data-stu-id="32144-229">String</span></span>|<span data-ttu-id="32144-230">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="32144-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32144-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="32144-231">Requirements</span></span>

|<span data-ttu-id="32144-232">Требование</span><span class="sxs-lookup"><span data-stu-id="32144-232">Requirement</span></span>| <span data-ttu-id="32144-233">Значение</span><span class="sxs-lookup"><span data-stu-id="32144-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="32144-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32144-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32144-235">1.1</span><span class="sxs-lookup"><span data-stu-id="32144-235">1.1</span></span>|
|[<span data-ttu-id="32144-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32144-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32144-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32144-237">Compose or Read</span></span>|
