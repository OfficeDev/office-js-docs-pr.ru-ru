---
title: Пространство имен Office — набор обязательных элементов 1,8
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 773a12d2f2b6c2d164b94d0b6b6c2dd0def90a41
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891182"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="cf18a-103">Office (набор требований для почтового ящика 1,8)</span><span class="sxs-lookup"><span data-stu-id="cf18a-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="cf18a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="cf18a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf18a-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf18a-106">Requirements</span></span>

|<span data-ttu-id="cf18a-107">Требование</span><span class="sxs-lookup"><span data-stu-id="cf18a-107">Requirement</span></span>| <span data-ttu-id="cf18a-108">Значение</span><span class="sxs-lookup"><span data-stu-id="cf18a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf18a-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cf18a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cf18a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-110">1.1</span></span>|
|[<span data-ttu-id="cf18a-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cf18a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cf18a-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cf18a-113">Properties</span><span class="sxs-lookup"><span data-stu-id="cf18a-113">Properties</span></span>

| <span data-ttu-id="cf18a-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="cf18a-114">Property</span></span> | <span data-ttu-id="cf18a-115">Способов</span><span class="sxs-lookup"><span data-stu-id="cf18a-115">Modes</span></span> | <span data-ttu-id="cf18a-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="cf18a-116">Return type</span></span> | <span data-ttu-id="cf18a-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="cf18a-117">Minimum</span></span><br><span data-ttu-id="cf18a-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="cf18a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cf18a-119">контекст</span><span class="sxs-lookup"><span data-stu-id="cf18a-119">context</span></span>](office.context.md) | <span data-ttu-id="cf18a-120">Создание</span><span class="sxs-lookup"><span data-stu-id="cf18a-120">Compose</span></span><br><span data-ttu-id="cf18a-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-121">Read</span></span> | [<span data-ttu-id="cf18a-122">Context</span><span class="sxs-lookup"><span data-stu-id="cf18a-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="cf18a-123">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cf18a-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="cf18a-124">Enumerations</span></span>

| <span data-ttu-id="cf18a-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="cf18a-125">Enumeration</span></span> | <span data-ttu-id="cf18a-126">Способов</span><span class="sxs-lookup"><span data-stu-id="cf18a-126">Modes</span></span> | <span data-ttu-id="cf18a-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="cf18a-127">Return type</span></span> | <span data-ttu-id="cf18a-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="cf18a-128">Minimum</span></span><br><span data-ttu-id="cf18a-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="cf18a-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cf18a-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cf18a-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cf18a-131">Создание</span><span class="sxs-lookup"><span data-stu-id="cf18a-131">Compose</span></span><br><span data-ttu-id="cf18a-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-132">Read</span></span> | <span data-ttu-id="cf18a-133">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-133">String</span></span> | [<span data-ttu-id="cf18a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cf18a-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cf18a-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cf18a-136">Создание</span><span class="sxs-lookup"><span data-stu-id="cf18a-136">Compose</span></span><br><span data-ttu-id="cf18a-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-137">Read</span></span> | <span data-ttu-id="cf18a-138">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-138">String</span></span> | [<span data-ttu-id="cf18a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cf18a-140">EventType</span><span class="sxs-lookup"><span data-stu-id="cf18a-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="cf18a-141">Создание</span><span class="sxs-lookup"><span data-stu-id="cf18a-141">Compose</span></span><br><span data-ttu-id="cf18a-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-142">Read</span></span> | <span data-ttu-id="cf18a-143">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-143">String</span></span> | [<span data-ttu-id="cf18a-144">1,5</span><span class="sxs-lookup"><span data-stu-id="cf18a-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="cf18a-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cf18a-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cf18a-146">Создание</span><span class="sxs-lookup"><span data-stu-id="cf18a-146">Compose</span></span><br><span data-ttu-id="cf18a-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-147">Read</span></span> | <span data-ttu-id="cf18a-148">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-148">String</span></span> | [<span data-ttu-id="cf18a-149">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cf18a-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="cf18a-150">Namespaces</span></span>

<span data-ttu-id="cf18a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="cf18a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cf18a-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="cf18a-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cf18a-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="cf18a-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="cf18a-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="cf18a-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cf18a-155">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-155">Type</span></span>

*   <span data-ttu-id="cf18a-156">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cf18a-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cf18a-157">Properties:</span></span>

|<span data-ttu-id="cf18a-158">Имя</span><span class="sxs-lookup"><span data-stu-id="cf18a-158">Name</span></span>| <span data-ttu-id="cf18a-159">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-159">Type</span></span>| <span data-ttu-id="cf18a-160">Описание</span><span class="sxs-lookup"><span data-stu-id="cf18a-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cf18a-161">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-161">String</span></span>|<span data-ttu-id="cf18a-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="cf18a-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cf18a-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="cf18a-163">String</span></span>|<span data-ttu-id="cf18a-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="cf18a-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf18a-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf18a-165">Requirements</span></span>

|<span data-ttu-id="cf18a-166">Требование</span><span class="sxs-lookup"><span data-stu-id="cf18a-166">Requirement</span></span>| <span data-ttu-id="cf18a-167">Значение</span><span class="sxs-lookup"><span data-stu-id="cf18a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf18a-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cf18a-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cf18a-169">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-169">1.1</span></span>|
|[<span data-ttu-id="cf18a-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cf18a-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cf18a-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cf18a-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="cf18a-172">CoercionType: String</span></span>

<span data-ttu-id="cf18a-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="cf18a-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cf18a-174">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-174">Type</span></span>

*   <span data-ttu-id="cf18a-175">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cf18a-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cf18a-176">Properties:</span></span>

|<span data-ttu-id="cf18a-177">Имя</span><span class="sxs-lookup"><span data-stu-id="cf18a-177">Name</span></span>| <span data-ttu-id="cf18a-178">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-178">Type</span></span>| <span data-ttu-id="cf18a-179">Описание</span><span class="sxs-lookup"><span data-stu-id="cf18a-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cf18a-180">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-180">String</span></span>|<span data-ttu-id="cf18a-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="cf18a-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cf18a-182">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-182">String</span></span>|<span data-ttu-id="cf18a-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="cf18a-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf18a-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf18a-184">Requirements</span></span>

|<span data-ttu-id="cf18a-185">Требование</span><span class="sxs-lookup"><span data-stu-id="cf18a-185">Requirement</span></span>| <span data-ttu-id="cf18a-186">Значение</span><span class="sxs-lookup"><span data-stu-id="cf18a-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf18a-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cf18a-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cf18a-188">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-188">1.1</span></span>|
|[<span data-ttu-id="cf18a-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cf18a-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cf18a-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="cf18a-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="cf18a-191">EventType: String</span></span>

<span data-ttu-id="cf18a-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="cf18a-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="cf18a-193">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-193">Type</span></span>

*   <span data-ttu-id="cf18a-194">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cf18a-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cf18a-195">Properties:</span></span>

| <span data-ttu-id="cf18a-196">Имя</span><span class="sxs-lookup"><span data-stu-id="cf18a-196">Name</span></span> | <span data-ttu-id="cf18a-197">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-197">Type</span></span> | <span data-ttu-id="cf18a-198">Описание</span><span class="sxs-lookup"><span data-stu-id="cf18a-198">Description</span></span> | <span data-ttu-id="cf18a-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="cf18a-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="cf18a-200">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-200">String</span></span> | <span data-ttu-id="cf18a-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="cf18a-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="cf18a-202">1.7</span><span class="sxs-lookup"><span data-stu-id="cf18a-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="cf18a-203">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-203">String</span></span> | <span data-ttu-id="cf18a-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="cf18a-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="cf18a-205">1.8</span><span class="sxs-lookup"><span data-stu-id="cf18a-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="cf18a-206">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-206">String</span></span> | <span data-ttu-id="cf18a-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="cf18a-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="cf18a-208">1.8</span><span class="sxs-lookup"><span data-stu-id="cf18a-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="cf18a-209">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-209">String</span></span> | <span data-ttu-id="cf18a-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="cf18a-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="cf18a-211">1.5</span><span class="sxs-lookup"><span data-stu-id="cf18a-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="cf18a-212">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-212">String</span></span> | <span data-ttu-id="cf18a-213">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="cf18a-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="cf18a-214">1.7</span><span class="sxs-lookup"><span data-stu-id="cf18a-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="cf18a-215">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-215">String</span></span> | <span data-ttu-id="cf18a-216">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="cf18a-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="cf18a-217">1.7</span><span class="sxs-lookup"><span data-stu-id="cf18a-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf18a-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf18a-218">Requirements</span></span>

|<span data-ttu-id="cf18a-219">Требование</span><span class="sxs-lookup"><span data-stu-id="cf18a-219">Requirement</span></span>| <span data-ttu-id="cf18a-220">Значение</span><span class="sxs-lookup"><span data-stu-id="cf18a-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf18a-221">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="cf18a-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cf18a-222">1.5</span><span class="sxs-lookup"><span data-stu-id="cf18a-222">1.5</span></span> |
|[<span data-ttu-id="cf18a-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cf18a-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cf18a-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cf18a-225">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="cf18a-225">SourceProperty: String</span></span>

<span data-ttu-id="cf18a-226">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="cf18a-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cf18a-227">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-227">Type</span></span>

*   <span data-ttu-id="cf18a-228">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cf18a-229">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cf18a-229">Properties:</span></span>

|<span data-ttu-id="cf18a-230">Имя</span><span class="sxs-lookup"><span data-stu-id="cf18a-230">Name</span></span>| <span data-ttu-id="cf18a-231">Тип</span><span class="sxs-lookup"><span data-stu-id="cf18a-231">Type</span></span>| <span data-ttu-id="cf18a-232">Описание</span><span class="sxs-lookup"><span data-stu-id="cf18a-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cf18a-233">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-233">String</span></span>|<span data-ttu-id="cf18a-234">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="cf18a-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cf18a-235">String</span><span class="sxs-lookup"><span data-stu-id="cf18a-235">String</span></span>|<span data-ttu-id="cf18a-236">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="cf18a-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf18a-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf18a-237">Requirements</span></span>

|<span data-ttu-id="cf18a-238">Требование</span><span class="sxs-lookup"><span data-stu-id="cf18a-238">Requirement</span></span>| <span data-ttu-id="cf18a-239">Значение</span><span class="sxs-lookup"><span data-stu-id="cf18a-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf18a-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cf18a-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cf18a-241">1.1</span><span class="sxs-lookup"><span data-stu-id="cf18a-241">1.1</span></span>|
|[<span data-ttu-id="cf18a-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cf18a-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cf18a-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cf18a-243">Compose or Read</span></span>|
