---
title: Пространство имен Office — набор обязательных элементов 1,9
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,9.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: e6a932c528dea692ff5fd7ea8d3e1454bb9a7e03
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628066"
---
# <a name="office-mailbox-requirement-set-19"></a><span data-ttu-id="e213e-103">Office (набор требований для почтового ящика 1,9)</span><span class="sxs-lookup"><span data-stu-id="e213e-103">Office (Mailbox requirement set 1.9)</span></span>

<span data-ttu-id="e213e-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e213e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e213e-106">Требования</span><span class="sxs-lookup"><span data-stu-id="e213e-106">Requirements</span></span>

|<span data-ttu-id="e213e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="e213e-107">Requirement</span></span>| <span data-ttu-id="e213e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="e213e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e213e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e213e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e213e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-110">1.1</span></span>|
|[<span data-ttu-id="e213e-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e213e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e213e-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e213e-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="e213e-113">Properties</span></span>

| <span data-ttu-id="e213e-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="e213e-114">Property</span></span> | <span data-ttu-id="e213e-115">Способов</span><span class="sxs-lookup"><span data-stu-id="e213e-115">Modes</span></span> | <span data-ttu-id="e213e-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e213e-116">Return type</span></span> | <span data-ttu-id="e213e-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="e213e-117">Minimum</span></span><br><span data-ttu-id="e213e-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="e213e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e213e-119">контекст</span><span class="sxs-lookup"><span data-stu-id="e213e-119">context</span></span>](office.context.md) | <span data-ttu-id="e213e-120">Создание</span><span class="sxs-lookup"><span data-stu-id="e213e-120">Compose</span></span><br><span data-ttu-id="e213e-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-121">Read</span></span> | [<span data-ttu-id="e213e-122">Context</span><span class="sxs-lookup"><span data-stu-id="e213e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e213e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="e213e-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="e213e-124">Enumerations</span></span>

| <span data-ttu-id="e213e-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="e213e-125">Enumeration</span></span> | <span data-ttu-id="e213e-126">Способов</span><span class="sxs-lookup"><span data-stu-id="e213e-126">Modes</span></span> | <span data-ttu-id="e213e-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e213e-127">Return type</span></span> | <span data-ttu-id="e213e-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="e213e-128">Minimum</span></span><br><span data-ttu-id="e213e-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="e213e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e213e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e213e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e213e-131">Создание</span><span class="sxs-lookup"><span data-stu-id="e213e-131">Compose</span></span><br><span data-ttu-id="e213e-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-132">Read</span></span> | <span data-ttu-id="e213e-133">String</span><span class="sxs-lookup"><span data-stu-id="e213e-133">String</span></span> | [<span data-ttu-id="e213e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e213e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e213e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e213e-136">Создание</span><span class="sxs-lookup"><span data-stu-id="e213e-136">Compose</span></span><br><span data-ttu-id="e213e-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-137">Read</span></span> | <span data-ttu-id="e213e-138">String</span><span class="sxs-lookup"><span data-stu-id="e213e-138">String</span></span> | [<span data-ttu-id="e213e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e213e-140">EventType</span><span class="sxs-lookup"><span data-stu-id="e213e-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e213e-141">Создание</span><span class="sxs-lookup"><span data-stu-id="e213e-141">Compose</span></span><br><span data-ttu-id="e213e-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-142">Read</span></span> | <span data-ttu-id="e213e-143">String</span><span class="sxs-lookup"><span data-stu-id="e213e-143">String</span></span> | [<span data-ttu-id="e213e-144">1,5</span><span class="sxs-lookup"><span data-stu-id="e213e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e213e-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e213e-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e213e-146">Создание</span><span class="sxs-lookup"><span data-stu-id="e213e-146">Compose</span></span><br><span data-ttu-id="e213e-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-147">Read</span></span> | <span data-ttu-id="e213e-148">String</span><span class="sxs-lookup"><span data-stu-id="e213e-148">String</span></span> | [<span data-ttu-id="e213e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="e213e-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="e213e-150">Namespaces</span></span>

<span data-ttu-id="e213e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="e213e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e213e-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="e213e-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e213e-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="e213e-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="e213e-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="e213e-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e213e-155">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-155">Type</span></span>

*   <span data-ttu-id="e213e-156">String</span><span class="sxs-lookup"><span data-stu-id="e213e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e213e-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e213e-157">Properties:</span></span>

|<span data-ttu-id="e213e-158">Имя</span><span class="sxs-lookup"><span data-stu-id="e213e-158">Name</span></span>| <span data-ttu-id="e213e-159">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-159">Type</span></span>| <span data-ttu-id="e213e-160">Описание</span><span class="sxs-lookup"><span data-stu-id="e213e-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e213e-161">String</span><span class="sxs-lookup"><span data-stu-id="e213e-161">String</span></span>|<span data-ttu-id="e213e-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="e213e-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e213e-163">String</span><span class="sxs-lookup"><span data-stu-id="e213e-163">String</span></span>|<span data-ttu-id="e213e-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="e213e-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e213e-165">Требования</span><span class="sxs-lookup"><span data-stu-id="e213e-165">Requirements</span></span>

|<span data-ttu-id="e213e-166">Требование</span><span class="sxs-lookup"><span data-stu-id="e213e-166">Requirement</span></span>| <span data-ttu-id="e213e-167">Значение</span><span class="sxs-lookup"><span data-stu-id="e213e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e213e-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e213e-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e213e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-169">1.1</span></span>|
|[<span data-ttu-id="e213e-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e213e-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e213e-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e213e-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="e213e-172">CoercionType: String</span></span>

<span data-ttu-id="e213e-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e213e-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e213e-174">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-174">Type</span></span>

*   <span data-ttu-id="e213e-175">String</span><span class="sxs-lookup"><span data-stu-id="e213e-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e213e-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e213e-176">Properties:</span></span>

|<span data-ttu-id="e213e-177">Имя</span><span class="sxs-lookup"><span data-stu-id="e213e-177">Name</span></span>| <span data-ttu-id="e213e-178">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-178">Type</span></span>| <span data-ttu-id="e213e-179">Описание</span><span class="sxs-lookup"><span data-stu-id="e213e-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e213e-180">String</span><span class="sxs-lookup"><span data-stu-id="e213e-180">String</span></span>|<span data-ttu-id="e213e-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="e213e-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e213e-182">String</span><span class="sxs-lookup"><span data-stu-id="e213e-182">String</span></span>|<span data-ttu-id="e213e-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="e213e-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e213e-184">Требования</span><span class="sxs-lookup"><span data-stu-id="e213e-184">Requirements</span></span>

|<span data-ttu-id="e213e-185">Требование</span><span class="sxs-lookup"><span data-stu-id="e213e-185">Requirement</span></span>| <span data-ttu-id="e213e-186">Значение</span><span class="sxs-lookup"><span data-stu-id="e213e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="e213e-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e213e-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e213e-188">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-188">1.1</span></span>|
|[<span data-ttu-id="e213e-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e213e-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e213e-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="e213e-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="e213e-191">EventType: String</span></span>

<span data-ttu-id="e213e-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="e213e-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e213e-193">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-193">Type</span></span>

*   <span data-ttu-id="e213e-194">String</span><span class="sxs-lookup"><span data-stu-id="e213e-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e213e-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e213e-195">Properties:</span></span>

| <span data-ttu-id="e213e-196">Имя</span><span class="sxs-lookup"><span data-stu-id="e213e-196">Name</span></span> | <span data-ttu-id="e213e-197">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-197">Type</span></span> | <span data-ttu-id="e213e-198">Описание</span><span class="sxs-lookup"><span data-stu-id="e213e-198">Description</span></span> | <span data-ttu-id="e213e-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="e213e-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="e213e-200">String</span><span class="sxs-lookup"><span data-stu-id="e213e-200">String</span></span> | <span data-ttu-id="e213e-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="e213e-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e213e-202">1.7</span><span class="sxs-lookup"><span data-stu-id="e213e-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="e213e-203">String</span><span class="sxs-lookup"><span data-stu-id="e213e-203">String</span></span> | <span data-ttu-id="e213e-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="e213e-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="e213e-205">1.8</span><span class="sxs-lookup"><span data-stu-id="e213e-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="e213e-206">String</span><span class="sxs-lookup"><span data-stu-id="e213e-206">String</span></span> | <span data-ttu-id="e213e-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="e213e-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="e213e-208">1.8</span><span class="sxs-lookup"><span data-stu-id="e213e-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="e213e-209">String</span><span class="sxs-lookup"><span data-stu-id="e213e-209">String</span></span> | <span data-ttu-id="e213e-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="e213e-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e213e-211">1.5</span><span class="sxs-lookup"><span data-stu-id="e213e-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e213e-212">String</span><span class="sxs-lookup"><span data-stu-id="e213e-212">String</span></span> | <span data-ttu-id="e213e-213">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="e213e-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e213e-214">1.7</span><span class="sxs-lookup"><span data-stu-id="e213e-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e213e-215">String</span><span class="sxs-lookup"><span data-stu-id="e213e-215">String</span></span> | <span data-ttu-id="e213e-216">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="e213e-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e213e-217">1.7</span><span class="sxs-lookup"><span data-stu-id="e213e-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e213e-218">Требования</span><span class="sxs-lookup"><span data-stu-id="e213e-218">Requirements</span></span>

|<span data-ttu-id="e213e-219">Требование</span><span class="sxs-lookup"><span data-stu-id="e213e-219">Requirement</span></span>| <span data-ttu-id="e213e-220">Значение</span><span class="sxs-lookup"><span data-stu-id="e213e-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="e213e-221">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e213e-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e213e-222">1.5</span><span class="sxs-lookup"><span data-stu-id="e213e-222">1.5</span></span> |
|[<span data-ttu-id="e213e-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e213e-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e213e-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-224">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e213e-225">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="e213e-225">SourceProperty: String</span></span>

<span data-ttu-id="e213e-226">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e213e-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e213e-227">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-227">Type</span></span>

*   <span data-ttu-id="e213e-228">String</span><span class="sxs-lookup"><span data-stu-id="e213e-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e213e-229">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e213e-229">Properties:</span></span>

|<span data-ttu-id="e213e-230">Имя</span><span class="sxs-lookup"><span data-stu-id="e213e-230">Name</span></span>| <span data-ttu-id="e213e-231">Тип</span><span class="sxs-lookup"><span data-stu-id="e213e-231">Type</span></span>| <span data-ttu-id="e213e-232">Описание</span><span class="sxs-lookup"><span data-stu-id="e213e-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e213e-233">String</span><span class="sxs-lookup"><span data-stu-id="e213e-233">String</span></span>|<span data-ttu-id="e213e-234">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="e213e-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e213e-235">String</span><span class="sxs-lookup"><span data-stu-id="e213e-235">String</span></span>|<span data-ttu-id="e213e-236">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="e213e-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e213e-237">Требования</span><span class="sxs-lookup"><span data-stu-id="e213e-237">Requirements</span></span>

|<span data-ttu-id="e213e-238">Требование</span><span class="sxs-lookup"><span data-stu-id="e213e-238">Requirement</span></span>| <span data-ttu-id="e213e-239">Значение</span><span class="sxs-lookup"><span data-stu-id="e213e-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="e213e-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e213e-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e213e-241">1.1</span><span class="sxs-lookup"><span data-stu-id="e213e-241">1.1</span></span>|
|[<span data-ttu-id="e213e-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e213e-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e213e-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e213e-243">Compose or Read</span></span>|
