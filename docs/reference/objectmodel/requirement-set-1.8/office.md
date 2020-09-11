---
title: Пространство имен Office — набор обязательных элементов 1,8
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: e0580cd1bb327c8673c46d3d0292aec9f2f1c971
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431523"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="76d33-103">Office (набор требований для почтового ящика 1,8)</span><span class="sxs-lookup"><span data-stu-id="76d33-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="76d33-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="76d33-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="76d33-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="76d33-106">Requirements</span></span>

|<span data-ttu-id="76d33-107">Требование</span><span class="sxs-lookup"><span data-stu-id="76d33-107">Requirement</span></span>| <span data-ttu-id="76d33-108">Значение</span><span class="sxs-lookup"><span data-stu-id="76d33-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="76d33-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76d33-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76d33-110">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-110">1.1</span></span>|
|[<span data-ttu-id="76d33-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76d33-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76d33-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="76d33-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="76d33-113">Properties</span></span>

| <span data-ttu-id="76d33-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="76d33-114">Property</span></span> | <span data-ttu-id="76d33-115">Способов</span><span class="sxs-lookup"><span data-stu-id="76d33-115">Modes</span></span> | <span data-ttu-id="76d33-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="76d33-116">Return type</span></span> | <span data-ttu-id="76d33-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="76d33-117">Minimum</span></span><br><span data-ttu-id="76d33-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="76d33-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="76d33-119">контекст</span><span class="sxs-lookup"><span data-stu-id="76d33-119">context</span></span>](office.context.md) | <span data-ttu-id="76d33-120">Создание</span><span class="sxs-lookup"><span data-stu-id="76d33-120">Compose</span></span><br><span data-ttu-id="76d33-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-121">Read</span></span> | [<span data-ttu-id="76d33-122">Context</span><span class="sxs-lookup"><span data-stu-id="76d33-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="76d33-123">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="76d33-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="76d33-124">Enumerations</span></span>

| <span data-ttu-id="76d33-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="76d33-125">Enumeration</span></span> | <span data-ttu-id="76d33-126">Способов</span><span class="sxs-lookup"><span data-stu-id="76d33-126">Modes</span></span> | <span data-ttu-id="76d33-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="76d33-127">Return type</span></span> | <span data-ttu-id="76d33-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="76d33-128">Minimum</span></span><br><span data-ttu-id="76d33-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="76d33-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="76d33-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="76d33-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="76d33-131">Создание</span><span class="sxs-lookup"><span data-stu-id="76d33-131">Compose</span></span><br><span data-ttu-id="76d33-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-132">Read</span></span> | <span data-ttu-id="76d33-133">String</span><span class="sxs-lookup"><span data-stu-id="76d33-133">String</span></span> | [<span data-ttu-id="76d33-134">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="76d33-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="76d33-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="76d33-136">Создание</span><span class="sxs-lookup"><span data-stu-id="76d33-136">Compose</span></span><br><span data-ttu-id="76d33-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-137">Read</span></span> | <span data-ttu-id="76d33-138">String</span><span class="sxs-lookup"><span data-stu-id="76d33-138">String</span></span> | [<span data-ttu-id="76d33-139">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="76d33-140">EventType</span><span class="sxs-lookup"><span data-stu-id="76d33-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="76d33-141">Создание</span><span class="sxs-lookup"><span data-stu-id="76d33-141">Compose</span></span><br><span data-ttu-id="76d33-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-142">Read</span></span> | <span data-ttu-id="76d33-143">String</span><span class="sxs-lookup"><span data-stu-id="76d33-143">String</span></span> | [<span data-ttu-id="76d33-144">1,5</span><span class="sxs-lookup"><span data-stu-id="76d33-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="76d33-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="76d33-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="76d33-146">Создание</span><span class="sxs-lookup"><span data-stu-id="76d33-146">Compose</span></span><br><span data-ttu-id="76d33-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-147">Read</span></span> | <span data-ttu-id="76d33-148">String</span><span class="sxs-lookup"><span data-stu-id="76d33-148">String</span></span> | [<span data-ttu-id="76d33-149">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="76d33-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="76d33-150">Namespaces</span></span>

<span data-ttu-id="76d33-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="76d33-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="76d33-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="76d33-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="76d33-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="76d33-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="76d33-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="76d33-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="76d33-155">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-155">Type</span></span>

*   <span data-ttu-id="76d33-156">String</span><span class="sxs-lookup"><span data-stu-id="76d33-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76d33-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76d33-157">Properties:</span></span>

|<span data-ttu-id="76d33-158">Имя</span><span class="sxs-lookup"><span data-stu-id="76d33-158">Name</span></span>| <span data-ttu-id="76d33-159">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-159">Type</span></span>| <span data-ttu-id="76d33-160">Описание</span><span class="sxs-lookup"><span data-stu-id="76d33-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="76d33-161">String</span><span class="sxs-lookup"><span data-stu-id="76d33-161">String</span></span>|<span data-ttu-id="76d33-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="76d33-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="76d33-163">String</span><span class="sxs-lookup"><span data-stu-id="76d33-163">String</span></span>|<span data-ttu-id="76d33-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="76d33-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76d33-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="76d33-165">Requirements</span></span>

|<span data-ttu-id="76d33-166">Требование</span><span class="sxs-lookup"><span data-stu-id="76d33-166">Requirement</span></span>| <span data-ttu-id="76d33-167">Значение</span><span class="sxs-lookup"><span data-stu-id="76d33-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="76d33-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76d33-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76d33-169">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-169">1.1</span></span>|
|[<span data-ttu-id="76d33-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76d33-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76d33-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="76d33-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="76d33-172">CoercionType: String</span></span>

<span data-ttu-id="76d33-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="76d33-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="76d33-174">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-174">Type</span></span>

*   <span data-ttu-id="76d33-175">String</span><span class="sxs-lookup"><span data-stu-id="76d33-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76d33-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76d33-176">Properties:</span></span>

|<span data-ttu-id="76d33-177">Имя</span><span class="sxs-lookup"><span data-stu-id="76d33-177">Name</span></span>| <span data-ttu-id="76d33-178">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-178">Type</span></span>| <span data-ttu-id="76d33-179">Описание</span><span class="sxs-lookup"><span data-stu-id="76d33-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="76d33-180">String</span><span class="sxs-lookup"><span data-stu-id="76d33-180">String</span></span>|<span data-ttu-id="76d33-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="76d33-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="76d33-182">String</span><span class="sxs-lookup"><span data-stu-id="76d33-182">String</span></span>|<span data-ttu-id="76d33-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="76d33-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76d33-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="76d33-184">Requirements</span></span>

|<span data-ttu-id="76d33-185">Требование</span><span class="sxs-lookup"><span data-stu-id="76d33-185">Requirement</span></span>| <span data-ttu-id="76d33-186">Значение</span><span class="sxs-lookup"><span data-stu-id="76d33-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="76d33-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76d33-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76d33-188">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-188">1.1</span></span>|
|[<span data-ttu-id="76d33-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76d33-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76d33-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="76d33-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="76d33-191">EventType: String</span></span>

<span data-ttu-id="76d33-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="76d33-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="76d33-193">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-193">Type</span></span>

*   <span data-ttu-id="76d33-194">String</span><span class="sxs-lookup"><span data-stu-id="76d33-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76d33-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76d33-195">Properties:</span></span>

| <span data-ttu-id="76d33-196">Имя</span><span class="sxs-lookup"><span data-stu-id="76d33-196">Name</span></span> | <span data-ttu-id="76d33-197">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-197">Type</span></span> | <span data-ttu-id="76d33-198">Описание</span><span class="sxs-lookup"><span data-stu-id="76d33-198">Description</span></span> | <span data-ttu-id="76d33-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="76d33-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="76d33-200">String</span><span class="sxs-lookup"><span data-stu-id="76d33-200">String</span></span> | <span data-ttu-id="76d33-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="76d33-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="76d33-202">1.7</span><span class="sxs-lookup"><span data-stu-id="76d33-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="76d33-203">String</span><span class="sxs-lookup"><span data-stu-id="76d33-203">String</span></span> | <span data-ttu-id="76d33-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="76d33-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="76d33-205">1.8</span><span class="sxs-lookup"><span data-stu-id="76d33-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="76d33-206">String</span><span class="sxs-lookup"><span data-stu-id="76d33-206">String</span></span> | <span data-ttu-id="76d33-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="76d33-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="76d33-208">1.8</span><span class="sxs-lookup"><span data-stu-id="76d33-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="76d33-209">String</span><span class="sxs-lookup"><span data-stu-id="76d33-209">String</span></span> | <span data-ttu-id="76d33-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="76d33-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="76d33-211">1.5</span><span class="sxs-lookup"><span data-stu-id="76d33-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="76d33-212">String</span><span class="sxs-lookup"><span data-stu-id="76d33-212">String</span></span> | <span data-ttu-id="76d33-213">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="76d33-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="76d33-214">1.7</span><span class="sxs-lookup"><span data-stu-id="76d33-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="76d33-215">String</span><span class="sxs-lookup"><span data-stu-id="76d33-215">String</span></span> | <span data-ttu-id="76d33-216">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="76d33-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="76d33-217">1.7</span><span class="sxs-lookup"><span data-stu-id="76d33-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="76d33-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="76d33-218">Requirements</span></span>

|<span data-ttu-id="76d33-219">Требование</span><span class="sxs-lookup"><span data-stu-id="76d33-219">Requirement</span></span>| <span data-ttu-id="76d33-220">Значение</span><span class="sxs-lookup"><span data-stu-id="76d33-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="76d33-221">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="76d33-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76d33-222">1.5</span><span class="sxs-lookup"><span data-stu-id="76d33-222">1.5</span></span> |
|[<span data-ttu-id="76d33-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76d33-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76d33-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="76d33-225">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="76d33-225">SourceProperty: String</span></span>

<span data-ttu-id="76d33-226">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="76d33-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="76d33-227">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-227">Type</span></span>

*   <span data-ttu-id="76d33-228">String</span><span class="sxs-lookup"><span data-stu-id="76d33-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="76d33-229">Свойства:</span><span class="sxs-lookup"><span data-stu-id="76d33-229">Properties:</span></span>

|<span data-ttu-id="76d33-230">Имя</span><span class="sxs-lookup"><span data-stu-id="76d33-230">Name</span></span>| <span data-ttu-id="76d33-231">Тип</span><span class="sxs-lookup"><span data-stu-id="76d33-231">Type</span></span>| <span data-ttu-id="76d33-232">Описание</span><span class="sxs-lookup"><span data-stu-id="76d33-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="76d33-233">String</span><span class="sxs-lookup"><span data-stu-id="76d33-233">String</span></span>|<span data-ttu-id="76d33-234">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="76d33-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="76d33-235">String</span><span class="sxs-lookup"><span data-stu-id="76d33-235">String</span></span>|<span data-ttu-id="76d33-236">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="76d33-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76d33-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="76d33-237">Requirements</span></span>

|<span data-ttu-id="76d33-238">Требование</span><span class="sxs-lookup"><span data-stu-id="76d33-238">Requirement</span></span>| <span data-ttu-id="76d33-239">Значение</span><span class="sxs-lookup"><span data-stu-id="76d33-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="76d33-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="76d33-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="76d33-241">1.1</span><span class="sxs-lookup"><span data-stu-id="76d33-241">1.1</span></span>|
|[<span data-ttu-id="76d33-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="76d33-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="76d33-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="76d33-243">Compose or Read</span></span>|
