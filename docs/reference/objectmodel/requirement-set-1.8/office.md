---
title: Пространство имен Office — набор обязательных элементов 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b23afd7b84dcd18e120f6aea4bd4fb0952791f1c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814168"
---
# <a name="office"></a><span data-ttu-id="0495c-102">Office</span><span class="sxs-lookup"><span data-stu-id="0495c-102">Office</span></span>

<span data-ttu-id="0495c-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0495c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0495c-105">Требования</span><span class="sxs-lookup"><span data-stu-id="0495c-105">Requirements</span></span>

|<span data-ttu-id="0495c-106">Требование</span><span class="sxs-lookup"><span data-stu-id="0495c-106">Requirement</span></span>| <span data-ttu-id="0495c-107">Значение</span><span class="sxs-lookup"><span data-stu-id="0495c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0495c-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0495c-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0495c-109">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-109">1.1</span></span>|
|[<span data-ttu-id="0495c-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0495c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0495c-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0495c-112">Properties</span><span class="sxs-lookup"><span data-stu-id="0495c-112">Properties</span></span>

| <span data-ttu-id="0495c-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="0495c-113">Property</span></span> | <span data-ttu-id="0495c-114">Способов</span><span class="sxs-lookup"><span data-stu-id="0495c-114">Modes</span></span> | <span data-ttu-id="0495c-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="0495c-115">Return type</span></span> | <span data-ttu-id="0495c-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="0495c-116">Minimum</span></span><br><span data-ttu-id="0495c-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="0495c-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0495c-118">контекст</span><span class="sxs-lookup"><span data-stu-id="0495c-118">context</span></span>](office.context.md) | <span data-ttu-id="0495c-119">Создание</span><span class="sxs-lookup"><span data-stu-id="0495c-119">Compose</span></span><br><span data-ttu-id="0495c-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-120">Read</span></span> | [<span data-ttu-id="0495c-121">Context</span><span class="sxs-lookup"><span data-stu-id="0495c-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="0495c-122">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="0495c-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="0495c-123">Enumerations</span></span>

| <span data-ttu-id="0495c-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="0495c-124">Enumeration</span></span> | <span data-ttu-id="0495c-125">Способов</span><span class="sxs-lookup"><span data-stu-id="0495c-125">Modes</span></span> | <span data-ttu-id="0495c-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="0495c-126">Return type</span></span> | <span data-ttu-id="0495c-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="0495c-127">Minimum</span></span><br><span data-ttu-id="0495c-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="0495c-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0495c-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0495c-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0495c-130">Создание</span><span class="sxs-lookup"><span data-stu-id="0495c-130">Compose</span></span><br><span data-ttu-id="0495c-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-131">Read</span></span> | <span data-ttu-id="0495c-132">String</span><span class="sxs-lookup"><span data-stu-id="0495c-132">String</span></span> | [<span data-ttu-id="0495c-133">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0495c-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0495c-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0495c-135">Создание</span><span class="sxs-lookup"><span data-stu-id="0495c-135">Compose</span></span><br><span data-ttu-id="0495c-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-136">Read</span></span> | <span data-ttu-id="0495c-137">String</span><span class="sxs-lookup"><span data-stu-id="0495c-137">String</span></span> | [<span data-ttu-id="0495c-138">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0495c-139">EventType</span><span class="sxs-lookup"><span data-stu-id="0495c-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0495c-140">Создание</span><span class="sxs-lookup"><span data-stu-id="0495c-140">Compose</span></span><br><span data-ttu-id="0495c-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-141">Read</span></span> | <span data-ttu-id="0495c-142">String</span><span class="sxs-lookup"><span data-stu-id="0495c-142">String</span></span> | [<span data-ttu-id="0495c-143">1,5</span><span class="sxs-lookup"><span data-stu-id="0495c-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="0495c-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0495c-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0495c-145">Создание</span><span class="sxs-lookup"><span data-stu-id="0495c-145">Compose</span></span><br><span data-ttu-id="0495c-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-146">Read</span></span> | <span data-ttu-id="0495c-147">String</span><span class="sxs-lookup"><span data-stu-id="0495c-147">String</span></span> | [<span data-ttu-id="0495c-148">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="0495c-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="0495c-149">Namespaces</span></span>

<span data-ttu-id="0495c-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="0495c-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="0495c-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="0495c-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="0495c-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="0495c-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="0495c-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="0495c-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0495c-154">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-154">Type</span></span>

*   <span data-ttu-id="0495c-155">String</span><span class="sxs-lookup"><span data-stu-id="0495c-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0495c-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0495c-156">Properties:</span></span>

|<span data-ttu-id="0495c-157">Имя</span><span class="sxs-lookup"><span data-stu-id="0495c-157">Name</span></span>| <span data-ttu-id="0495c-158">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-158">Type</span></span>| <span data-ttu-id="0495c-159">Описание</span><span class="sxs-lookup"><span data-stu-id="0495c-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0495c-160">String</span><span class="sxs-lookup"><span data-stu-id="0495c-160">String</span></span>|<span data-ttu-id="0495c-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="0495c-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0495c-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="0495c-162">String</span></span>|<span data-ttu-id="0495c-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="0495c-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0495c-164">Требования</span><span class="sxs-lookup"><span data-stu-id="0495c-164">Requirements</span></span>

|<span data-ttu-id="0495c-165">Требование</span><span class="sxs-lookup"><span data-stu-id="0495c-165">Requirement</span></span>| <span data-ttu-id="0495c-166">Значение</span><span class="sxs-lookup"><span data-stu-id="0495c-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="0495c-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0495c-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0495c-168">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-168">1.1</span></span>|
|[<span data-ttu-id="0495c-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0495c-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0495c-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="0495c-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="0495c-171">CoercionType: String</span></span>

<span data-ttu-id="0495c-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0495c-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0495c-173">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-173">Type</span></span>

*   <span data-ttu-id="0495c-174">String</span><span class="sxs-lookup"><span data-stu-id="0495c-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0495c-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0495c-175">Properties:</span></span>

|<span data-ttu-id="0495c-176">Имя</span><span class="sxs-lookup"><span data-stu-id="0495c-176">Name</span></span>| <span data-ttu-id="0495c-177">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-177">Type</span></span>| <span data-ttu-id="0495c-178">Описание</span><span class="sxs-lookup"><span data-stu-id="0495c-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0495c-179">String</span><span class="sxs-lookup"><span data-stu-id="0495c-179">String</span></span>|<span data-ttu-id="0495c-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="0495c-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0495c-181">String</span><span class="sxs-lookup"><span data-stu-id="0495c-181">String</span></span>|<span data-ttu-id="0495c-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="0495c-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0495c-183">Требования</span><span class="sxs-lookup"><span data-stu-id="0495c-183">Requirements</span></span>

|<span data-ttu-id="0495c-184">Требование</span><span class="sxs-lookup"><span data-stu-id="0495c-184">Requirement</span></span>| <span data-ttu-id="0495c-185">Значение</span><span class="sxs-lookup"><span data-stu-id="0495c-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="0495c-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0495c-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0495c-187">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-187">1.1</span></span>|
|[<span data-ttu-id="0495c-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0495c-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0495c-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="0495c-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="0495c-190">EventType: String</span></span>

<span data-ttu-id="0495c-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="0495c-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0495c-192">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-192">Type</span></span>

*   <span data-ttu-id="0495c-193">String</span><span class="sxs-lookup"><span data-stu-id="0495c-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0495c-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0495c-194">Properties:</span></span>

| <span data-ttu-id="0495c-195">Имя</span><span class="sxs-lookup"><span data-stu-id="0495c-195">Name</span></span> | <span data-ttu-id="0495c-196">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-196">Type</span></span> | <span data-ttu-id="0495c-197">Описание</span><span class="sxs-lookup"><span data-stu-id="0495c-197">Description</span></span> | <span data-ttu-id="0495c-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="0495c-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="0495c-199">String</span><span class="sxs-lookup"><span data-stu-id="0495c-199">String</span></span> | <span data-ttu-id="0495c-200">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="0495c-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="0495c-201">1.7</span><span class="sxs-lookup"><span data-stu-id="0495c-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="0495c-202">String</span><span class="sxs-lookup"><span data-stu-id="0495c-202">String</span></span> | <span data-ttu-id="0495c-203">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="0495c-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="0495c-204">1.8</span><span class="sxs-lookup"><span data-stu-id="0495c-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="0495c-205">String</span><span class="sxs-lookup"><span data-stu-id="0495c-205">String</span></span> | <span data-ttu-id="0495c-206">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="0495c-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="0495c-207">1.8</span><span class="sxs-lookup"><span data-stu-id="0495c-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="0495c-208">String</span><span class="sxs-lookup"><span data-stu-id="0495c-208">String</span></span> | <span data-ttu-id="0495c-209">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="0495c-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="0495c-210">1.5</span><span class="sxs-lookup"><span data-stu-id="0495c-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="0495c-211">String</span><span class="sxs-lookup"><span data-stu-id="0495c-211">String</span></span> | <span data-ttu-id="0495c-212">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="0495c-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="0495c-213">1.7</span><span class="sxs-lookup"><span data-stu-id="0495c-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="0495c-214">String</span><span class="sxs-lookup"><span data-stu-id="0495c-214">String</span></span> | <span data-ttu-id="0495c-215">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="0495c-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="0495c-216">1.7</span><span class="sxs-lookup"><span data-stu-id="0495c-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0495c-217">Требования</span><span class="sxs-lookup"><span data-stu-id="0495c-217">Requirements</span></span>

|<span data-ttu-id="0495c-218">Требование</span><span class="sxs-lookup"><span data-stu-id="0495c-218">Requirement</span></span>| <span data-ttu-id="0495c-219">Значение</span><span class="sxs-lookup"><span data-stu-id="0495c-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="0495c-220">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0495c-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0495c-221">1.5</span><span class="sxs-lookup"><span data-stu-id="0495c-221">1.5</span></span> |
|[<span data-ttu-id="0495c-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0495c-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0495c-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="0495c-224">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="0495c-224">SourceProperty: String</span></span>

<span data-ttu-id="0495c-225">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0495c-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0495c-226">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-226">Type</span></span>

*   <span data-ttu-id="0495c-227">String</span><span class="sxs-lookup"><span data-stu-id="0495c-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0495c-228">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0495c-228">Properties:</span></span>

|<span data-ttu-id="0495c-229">Имя</span><span class="sxs-lookup"><span data-stu-id="0495c-229">Name</span></span>| <span data-ttu-id="0495c-230">Тип</span><span class="sxs-lookup"><span data-stu-id="0495c-230">Type</span></span>| <span data-ttu-id="0495c-231">Описание</span><span class="sxs-lookup"><span data-stu-id="0495c-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0495c-232">String</span><span class="sxs-lookup"><span data-stu-id="0495c-232">String</span></span>|<span data-ttu-id="0495c-233">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="0495c-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0495c-234">String</span><span class="sxs-lookup"><span data-stu-id="0495c-234">String</span></span>|<span data-ttu-id="0495c-235">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="0495c-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0495c-236">Требования</span><span class="sxs-lookup"><span data-stu-id="0495c-236">Requirements</span></span>

|<span data-ttu-id="0495c-237">Требование</span><span class="sxs-lookup"><span data-stu-id="0495c-237">Requirement</span></span>| <span data-ttu-id="0495c-238">Значение</span><span class="sxs-lookup"><span data-stu-id="0495c-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="0495c-239">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0495c-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0495c-240">1.1</span><span class="sxs-lookup"><span data-stu-id="0495c-240">1.1</span></span>|
|[<span data-ttu-id="0495c-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0495c-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0495c-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0495c-242">Compose or Read</span></span>|
