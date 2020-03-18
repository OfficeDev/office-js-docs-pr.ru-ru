---
title: Пространство имен Office — набор обязательных элементов 1,8
description: Пространство имен Office предоставляет общие интерфейсы для надстроек Outlook Office (набор требований 1,8)
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0bbe212b0b8e5dc1348cb5cdc03509c44a716d1a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717505"
---
# <a name="office"></a><span data-ttu-id="32080-103">Office</span><span class="sxs-lookup"><span data-stu-id="32080-103">Office</span></span>

<span data-ttu-id="32080-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="32080-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="32080-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="32080-106">Requirements</span></span>

|<span data-ttu-id="32080-107">Требование</span><span class="sxs-lookup"><span data-stu-id="32080-107">Requirement</span></span>| <span data-ttu-id="32080-108">Значение</span><span class="sxs-lookup"><span data-stu-id="32080-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="32080-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32080-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32080-110">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-110">1.1</span></span>|
|[<span data-ttu-id="32080-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32080-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32080-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32080-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="32080-113">Properties</span><span class="sxs-lookup"><span data-stu-id="32080-113">Properties</span></span>

| <span data-ttu-id="32080-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="32080-114">Property</span></span> | <span data-ttu-id="32080-115">Способов</span><span class="sxs-lookup"><span data-stu-id="32080-115">Modes</span></span> | <span data-ttu-id="32080-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="32080-116">Return type</span></span> | <span data-ttu-id="32080-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="32080-117">Minimum</span></span><br><span data-ttu-id="32080-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="32080-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="32080-119">контекст</span><span class="sxs-lookup"><span data-stu-id="32080-119">context</span></span>](office.context.md) | <span data-ttu-id="32080-120">Создание</span><span class="sxs-lookup"><span data-stu-id="32080-120">Compose</span></span><br><span data-ttu-id="32080-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="32080-121">Read</span></span> | [<span data-ttu-id="32080-122">Context</span><span class="sxs-lookup"><span data-stu-id="32080-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="32080-123">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="32080-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="32080-124">Enumerations</span></span>

| <span data-ttu-id="32080-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="32080-125">Enumeration</span></span> | <span data-ttu-id="32080-126">Способов</span><span class="sxs-lookup"><span data-stu-id="32080-126">Modes</span></span> | <span data-ttu-id="32080-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="32080-127">Return type</span></span> | <span data-ttu-id="32080-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="32080-128">Minimum</span></span><br><span data-ttu-id="32080-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="32080-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="32080-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="32080-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="32080-131">Создание</span><span class="sxs-lookup"><span data-stu-id="32080-131">Compose</span></span><br><span data-ttu-id="32080-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="32080-132">Read</span></span> | <span data-ttu-id="32080-133">String</span><span class="sxs-lookup"><span data-stu-id="32080-133">String</span></span> | [<span data-ttu-id="32080-134">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32080-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="32080-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="32080-136">Создание</span><span class="sxs-lookup"><span data-stu-id="32080-136">Compose</span></span><br><span data-ttu-id="32080-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="32080-137">Read</span></span> | <span data-ttu-id="32080-138">String</span><span class="sxs-lookup"><span data-stu-id="32080-138">String</span></span> | [<span data-ttu-id="32080-139">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32080-140">EventType</span><span class="sxs-lookup"><span data-stu-id="32080-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="32080-141">Создание</span><span class="sxs-lookup"><span data-stu-id="32080-141">Compose</span></span><br><span data-ttu-id="32080-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="32080-142">Read</span></span> | <span data-ttu-id="32080-143">String</span><span class="sxs-lookup"><span data-stu-id="32080-143">String</span></span> | [<span data-ttu-id="32080-144">1,5</span><span class="sxs-lookup"><span data-stu-id="32080-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="32080-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="32080-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="32080-146">Создание</span><span class="sxs-lookup"><span data-stu-id="32080-146">Compose</span></span><br><span data-ttu-id="32080-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="32080-147">Read</span></span> | <span data-ttu-id="32080-148">String</span><span class="sxs-lookup"><span data-stu-id="32080-148">String</span></span> | [<span data-ttu-id="32080-149">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="32080-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="32080-150">Namespaces</span></span>

<span data-ttu-id="32080-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="32080-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="32080-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="32080-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="32080-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="32080-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="32080-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="32080-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="32080-155">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-155">Type</span></span>

*   <span data-ttu-id="32080-156">String</span><span class="sxs-lookup"><span data-stu-id="32080-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32080-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32080-157">Properties:</span></span>

|<span data-ttu-id="32080-158">Имя</span><span class="sxs-lookup"><span data-stu-id="32080-158">Name</span></span>| <span data-ttu-id="32080-159">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-159">Type</span></span>| <span data-ttu-id="32080-160">Описание</span><span class="sxs-lookup"><span data-stu-id="32080-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="32080-161">String</span><span class="sxs-lookup"><span data-stu-id="32080-161">String</span></span>|<span data-ttu-id="32080-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="32080-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="32080-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="32080-163">String</span></span>|<span data-ttu-id="32080-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="32080-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32080-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="32080-165">Requirements</span></span>

|<span data-ttu-id="32080-166">Требование</span><span class="sxs-lookup"><span data-stu-id="32080-166">Requirement</span></span>| <span data-ttu-id="32080-167">Значение</span><span class="sxs-lookup"><span data-stu-id="32080-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="32080-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32080-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32080-169">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-169">1.1</span></span>|
|[<span data-ttu-id="32080-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32080-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32080-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32080-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="32080-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="32080-172">CoercionType: String</span></span>

<span data-ttu-id="32080-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="32080-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="32080-174">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-174">Type</span></span>

*   <span data-ttu-id="32080-175">String</span><span class="sxs-lookup"><span data-stu-id="32080-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32080-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32080-176">Properties:</span></span>

|<span data-ttu-id="32080-177">Имя</span><span class="sxs-lookup"><span data-stu-id="32080-177">Name</span></span>| <span data-ttu-id="32080-178">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-178">Type</span></span>| <span data-ttu-id="32080-179">Описание</span><span class="sxs-lookup"><span data-stu-id="32080-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="32080-180">String</span><span class="sxs-lookup"><span data-stu-id="32080-180">String</span></span>|<span data-ttu-id="32080-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="32080-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="32080-182">String</span><span class="sxs-lookup"><span data-stu-id="32080-182">String</span></span>|<span data-ttu-id="32080-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="32080-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32080-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="32080-184">Requirements</span></span>

|<span data-ttu-id="32080-185">Требование</span><span class="sxs-lookup"><span data-stu-id="32080-185">Requirement</span></span>| <span data-ttu-id="32080-186">Значение</span><span class="sxs-lookup"><span data-stu-id="32080-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="32080-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32080-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32080-188">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-188">1.1</span></span>|
|[<span data-ttu-id="32080-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32080-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32080-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32080-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="32080-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="32080-191">EventType: String</span></span>

<span data-ttu-id="32080-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="32080-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="32080-193">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-193">Type</span></span>

*   <span data-ttu-id="32080-194">String</span><span class="sxs-lookup"><span data-stu-id="32080-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32080-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32080-195">Properties:</span></span>

| <span data-ttu-id="32080-196">Имя</span><span class="sxs-lookup"><span data-stu-id="32080-196">Name</span></span> | <span data-ttu-id="32080-197">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-197">Type</span></span> | <span data-ttu-id="32080-198">Описание</span><span class="sxs-lookup"><span data-stu-id="32080-198">Description</span></span> | <span data-ttu-id="32080-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="32080-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="32080-200">String</span><span class="sxs-lookup"><span data-stu-id="32080-200">String</span></span> | <span data-ttu-id="32080-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="32080-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="32080-202">1.7</span><span class="sxs-lookup"><span data-stu-id="32080-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="32080-203">String</span><span class="sxs-lookup"><span data-stu-id="32080-203">String</span></span> | <span data-ttu-id="32080-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="32080-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="32080-205">1.8</span><span class="sxs-lookup"><span data-stu-id="32080-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="32080-206">String</span><span class="sxs-lookup"><span data-stu-id="32080-206">String</span></span> | <span data-ttu-id="32080-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="32080-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="32080-208">1.8</span><span class="sxs-lookup"><span data-stu-id="32080-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="32080-209">String</span><span class="sxs-lookup"><span data-stu-id="32080-209">String</span></span> | <span data-ttu-id="32080-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="32080-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="32080-211">1.5</span><span class="sxs-lookup"><span data-stu-id="32080-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="32080-212">String</span><span class="sxs-lookup"><span data-stu-id="32080-212">String</span></span> | <span data-ttu-id="32080-213">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="32080-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="32080-214">1.7</span><span class="sxs-lookup"><span data-stu-id="32080-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="32080-215">String</span><span class="sxs-lookup"><span data-stu-id="32080-215">String</span></span> | <span data-ttu-id="32080-216">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="32080-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="32080-217">1.7</span><span class="sxs-lookup"><span data-stu-id="32080-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="32080-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="32080-218">Requirements</span></span>

|<span data-ttu-id="32080-219">Требование</span><span class="sxs-lookup"><span data-stu-id="32080-219">Requirement</span></span>| <span data-ttu-id="32080-220">Значение</span><span class="sxs-lookup"><span data-stu-id="32080-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="32080-221">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="32080-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32080-222">1.5</span><span class="sxs-lookup"><span data-stu-id="32080-222">1.5</span></span> |
|[<span data-ttu-id="32080-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32080-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32080-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32080-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="32080-225">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="32080-225">SourceProperty: String</span></span>

<span data-ttu-id="32080-226">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="32080-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="32080-227">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-227">Type</span></span>

*   <span data-ttu-id="32080-228">String</span><span class="sxs-lookup"><span data-stu-id="32080-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32080-229">Свойства:</span><span class="sxs-lookup"><span data-stu-id="32080-229">Properties:</span></span>

|<span data-ttu-id="32080-230">Имя</span><span class="sxs-lookup"><span data-stu-id="32080-230">Name</span></span>| <span data-ttu-id="32080-231">Тип</span><span class="sxs-lookup"><span data-stu-id="32080-231">Type</span></span>| <span data-ttu-id="32080-232">Описание</span><span class="sxs-lookup"><span data-stu-id="32080-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="32080-233">String</span><span class="sxs-lookup"><span data-stu-id="32080-233">String</span></span>|<span data-ttu-id="32080-234">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="32080-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="32080-235">String</span><span class="sxs-lookup"><span data-stu-id="32080-235">String</span></span>|<span data-ttu-id="32080-236">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="32080-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32080-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="32080-237">Requirements</span></span>

|<span data-ttu-id="32080-238">Требование</span><span class="sxs-lookup"><span data-stu-id="32080-238">Requirement</span></span>| <span data-ttu-id="32080-239">Значение</span><span class="sxs-lookup"><span data-stu-id="32080-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="32080-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="32080-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32080-241">1.1</span><span class="sxs-lookup"><span data-stu-id="32080-241">1.1</span></span>|
|[<span data-ttu-id="32080-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="32080-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32080-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="32080-243">Compose or Read</span></span>|
