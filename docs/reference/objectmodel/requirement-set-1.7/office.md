---
title: Пространство имен Office — набор обязательных элементов 1,7
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 718de46689fc2fcb52ad455763581ecab06a4c39
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612201"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="d9829-103">Office (набор требований для почтового ящика 1,7)</span><span class="sxs-lookup"><span data-stu-id="d9829-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="d9829-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d9829-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9829-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="d9829-106">Requirements</span></span>

|<span data-ttu-id="d9829-107">Требование</span><span class="sxs-lookup"><span data-stu-id="d9829-107">Requirement</span></span>| <span data-ttu-id="d9829-108">Значение</span><span class="sxs-lookup"><span data-stu-id="d9829-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9829-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d9829-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d9829-110">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-110">1.1</span></span>|
|[<span data-ttu-id="d9829-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9829-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d9829-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9829-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d9829-113">Properties</span><span class="sxs-lookup"><span data-stu-id="d9829-113">Properties</span></span>

| <span data-ttu-id="d9829-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="d9829-114">Property</span></span> | <span data-ttu-id="d9829-115">Способов</span><span class="sxs-lookup"><span data-stu-id="d9829-115">Modes</span></span> | <span data-ttu-id="d9829-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="d9829-116">Return type</span></span> | <span data-ttu-id="d9829-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="d9829-117">Minimum</span></span><br><span data-ttu-id="d9829-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="d9829-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d9829-119">контекст</span><span class="sxs-lookup"><span data-stu-id="d9829-119">context</span></span>](office.context.md) | <span data-ttu-id="d9829-120">Создание</span><span class="sxs-lookup"><span data-stu-id="d9829-120">Compose</span></span><br><span data-ttu-id="d9829-121">Read</span><span class="sxs-lookup"><span data-stu-id="d9829-121">Read</span></span> | [<span data-ttu-id="d9829-122">Context</span><span class="sxs-lookup"><span data-stu-id="d9829-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="d9829-123">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="d9829-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="d9829-124">Enumerations</span></span>

| <span data-ttu-id="d9829-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="d9829-125">Enumeration</span></span> | <span data-ttu-id="d9829-126">Способов</span><span class="sxs-lookup"><span data-stu-id="d9829-126">Modes</span></span> | <span data-ttu-id="d9829-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="d9829-127">Return type</span></span> | <span data-ttu-id="d9829-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="d9829-128">Minimum</span></span><br><span data-ttu-id="d9829-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="d9829-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d9829-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d9829-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d9829-131">Создание</span><span class="sxs-lookup"><span data-stu-id="d9829-131">Compose</span></span><br><span data-ttu-id="d9829-132">Read</span><span class="sxs-lookup"><span data-stu-id="d9829-132">Read</span></span> | <span data-ttu-id="d9829-133">String</span><span class="sxs-lookup"><span data-stu-id="d9829-133">String</span></span> | [<span data-ttu-id="d9829-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d9829-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d9829-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d9829-136">Создание</span><span class="sxs-lookup"><span data-stu-id="d9829-136">Compose</span></span><br><span data-ttu-id="d9829-137">Read</span><span class="sxs-lookup"><span data-stu-id="d9829-137">Read</span></span> | <span data-ttu-id="d9829-138">String</span><span class="sxs-lookup"><span data-stu-id="d9829-138">String</span></span> | [<span data-ttu-id="d9829-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d9829-140">EventType</span><span class="sxs-lookup"><span data-stu-id="d9829-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d9829-141">Создание</span><span class="sxs-lookup"><span data-stu-id="d9829-141">Compose</span></span><br><span data-ttu-id="d9829-142">Read</span><span class="sxs-lookup"><span data-stu-id="d9829-142">Read</span></span> | <span data-ttu-id="d9829-143">String</span><span class="sxs-lookup"><span data-stu-id="d9829-143">String</span></span> | [<span data-ttu-id="d9829-144">1,5</span><span class="sxs-lookup"><span data-stu-id="d9829-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="d9829-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d9829-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d9829-146">Создание</span><span class="sxs-lookup"><span data-stu-id="d9829-146">Compose</span></span><br><span data-ttu-id="d9829-147">Read</span><span class="sxs-lookup"><span data-stu-id="d9829-147">Read</span></span> | <span data-ttu-id="d9829-148">String</span><span class="sxs-lookup"><span data-stu-id="d9829-148">String</span></span> | [<span data-ttu-id="d9829-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="d9829-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d9829-150">Namespaces</span></span>

<span data-ttu-id="d9829-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="d9829-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d9829-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="d9829-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d9829-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="d9829-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="d9829-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9829-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d9829-155">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-155">Type</span></span>

*   <span data-ttu-id="d9829-156">String</span><span class="sxs-lookup"><span data-stu-id="d9829-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9829-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d9829-157">Properties:</span></span>

|<span data-ttu-id="d9829-158">Имя</span><span class="sxs-lookup"><span data-stu-id="d9829-158">Name</span></span>| <span data-ttu-id="d9829-159">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-159">Type</span></span>| <span data-ttu-id="d9829-160">Описание</span><span class="sxs-lookup"><span data-stu-id="d9829-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d9829-161">String</span><span class="sxs-lookup"><span data-stu-id="d9829-161">String</span></span>|<span data-ttu-id="d9829-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="d9829-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d9829-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="d9829-163">String</span></span>|<span data-ttu-id="d9829-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="d9829-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9829-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="d9829-165">Requirements</span></span>

|<span data-ttu-id="d9829-166">Требование</span><span class="sxs-lookup"><span data-stu-id="d9829-166">Requirement</span></span>| <span data-ttu-id="d9829-167">Значение</span><span class="sxs-lookup"><span data-stu-id="d9829-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9829-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d9829-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d9829-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-169">1.1</span></span>|
|[<span data-ttu-id="d9829-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9829-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d9829-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9829-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d9829-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="d9829-172">CoercionType: String</span></span>

<span data-ttu-id="d9829-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d9829-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d9829-174">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-174">Type</span></span>

*   <span data-ttu-id="d9829-175">String</span><span class="sxs-lookup"><span data-stu-id="d9829-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9829-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d9829-176">Properties:</span></span>

|<span data-ttu-id="d9829-177">Имя</span><span class="sxs-lookup"><span data-stu-id="d9829-177">Name</span></span>| <span data-ttu-id="d9829-178">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-178">Type</span></span>| <span data-ttu-id="d9829-179">Описание</span><span class="sxs-lookup"><span data-stu-id="d9829-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d9829-180">String</span><span class="sxs-lookup"><span data-stu-id="d9829-180">String</span></span>|<span data-ttu-id="d9829-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d9829-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d9829-182">String</span><span class="sxs-lookup"><span data-stu-id="d9829-182">String</span></span>|<span data-ttu-id="d9829-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="d9829-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9829-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="d9829-184">Requirements</span></span>

|<span data-ttu-id="d9829-185">Требование</span><span class="sxs-lookup"><span data-stu-id="d9829-185">Requirement</span></span>| <span data-ttu-id="d9829-186">Значение</span><span class="sxs-lookup"><span data-stu-id="d9829-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9829-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d9829-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d9829-188">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-188">1.1</span></span>|
|[<span data-ttu-id="d9829-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9829-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d9829-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9829-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="d9829-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="d9829-191">EventType: String</span></span>

<span data-ttu-id="d9829-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="d9829-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d9829-193">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-193">Type</span></span>

*   <span data-ttu-id="d9829-194">String</span><span class="sxs-lookup"><span data-stu-id="d9829-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9829-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d9829-195">Properties:</span></span>

| <span data-ttu-id="d9829-196">Имя</span><span class="sxs-lookup"><span data-stu-id="d9829-196">Name</span></span> | <span data-ttu-id="d9829-197">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-197">Type</span></span> | <span data-ttu-id="d9829-198">Описание</span><span class="sxs-lookup"><span data-stu-id="d9829-198">Description</span></span> | <span data-ttu-id="d9829-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="d9829-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="d9829-200">String</span><span class="sxs-lookup"><span data-stu-id="d9829-200">String</span></span> | <span data-ttu-id="d9829-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="d9829-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="d9829-202">1.7</span><span class="sxs-lookup"><span data-stu-id="d9829-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="d9829-203">String</span><span class="sxs-lookup"><span data-stu-id="d9829-203">String</span></span> | <span data-ttu-id="d9829-204">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="d9829-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d9829-205">1.5</span><span class="sxs-lookup"><span data-stu-id="d9829-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="d9829-206">String</span><span class="sxs-lookup"><span data-stu-id="d9829-206">String</span></span> | <span data-ttu-id="d9829-207">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="d9829-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="d9829-208">1.7</span><span class="sxs-lookup"><span data-stu-id="d9829-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="d9829-209">String</span><span class="sxs-lookup"><span data-stu-id="d9829-209">String</span></span> | <span data-ttu-id="d9829-210">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="d9829-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="d9829-211">1.7</span><span class="sxs-lookup"><span data-stu-id="d9829-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9829-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="d9829-212">Requirements</span></span>

|<span data-ttu-id="d9829-213">Требование</span><span class="sxs-lookup"><span data-stu-id="d9829-213">Requirement</span></span>| <span data-ttu-id="d9829-214">Значение</span><span class="sxs-lookup"><span data-stu-id="d9829-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9829-215">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d9829-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d9829-216">1.5</span><span class="sxs-lookup"><span data-stu-id="d9829-216">1.5</span></span> |
|[<span data-ttu-id="d9829-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9829-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d9829-218">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9829-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d9829-219">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="d9829-219">SourceProperty: String</span></span>

<span data-ttu-id="d9829-220">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d9829-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d9829-221">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-221">Type</span></span>

*   <span data-ttu-id="d9829-222">String</span><span class="sxs-lookup"><span data-stu-id="d9829-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9829-223">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d9829-223">Properties:</span></span>

|<span data-ttu-id="d9829-224">Имя</span><span class="sxs-lookup"><span data-stu-id="d9829-224">Name</span></span>| <span data-ttu-id="d9829-225">Тип</span><span class="sxs-lookup"><span data-stu-id="d9829-225">Type</span></span>| <span data-ttu-id="d9829-226">Описание</span><span class="sxs-lookup"><span data-stu-id="d9829-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d9829-227">String</span><span class="sxs-lookup"><span data-stu-id="d9829-227">String</span></span>|<span data-ttu-id="d9829-228">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9829-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d9829-229">String</span><span class="sxs-lookup"><span data-stu-id="d9829-229">String</span></span>|<span data-ttu-id="d9829-230">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9829-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9829-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="d9829-231">Requirements</span></span>

|<span data-ttu-id="d9829-232">Требование</span><span class="sxs-lookup"><span data-stu-id="d9829-232">Requirement</span></span>| <span data-ttu-id="d9829-233">Значение</span><span class="sxs-lookup"><span data-stu-id="d9829-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9829-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d9829-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d9829-235">1.1</span><span class="sxs-lookup"><span data-stu-id="d9829-235">1.1</span></span>|
|[<span data-ttu-id="d9829-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9829-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d9829-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9829-237">Compose or Read</span></span>|
