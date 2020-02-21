---
title: Пространство имен Office — набор обязательных элементов 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: c5c431f7a958f1c2a956f36e90ad0f3a205c6669
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163628"
---
# <a name="office"></a><span data-ttu-id="70730-102">Office</span><span class="sxs-lookup"><span data-stu-id="70730-102">Office</span></span>

<span data-ttu-id="70730-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="70730-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="70730-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="70730-105">Requirements</span></span>

|<span data-ttu-id="70730-106">Требование</span><span class="sxs-lookup"><span data-stu-id="70730-106">Requirement</span></span>| <span data-ttu-id="70730-107">Значение</span><span class="sxs-lookup"><span data-stu-id="70730-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="70730-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="70730-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="70730-109">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-109">1.1</span></span>|
|[<span data-ttu-id="70730-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="70730-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="70730-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="70730-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="70730-112">Properties</span><span class="sxs-lookup"><span data-stu-id="70730-112">Properties</span></span>

| <span data-ttu-id="70730-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="70730-113">Property</span></span> | <span data-ttu-id="70730-114">Способов</span><span class="sxs-lookup"><span data-stu-id="70730-114">Modes</span></span> | <span data-ttu-id="70730-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="70730-115">Return type</span></span> | <span data-ttu-id="70730-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="70730-116">Minimum</span></span><br><span data-ttu-id="70730-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="70730-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="70730-118">контекст</span><span class="sxs-lookup"><span data-stu-id="70730-118">context</span></span>](office.context.md) | <span data-ttu-id="70730-119">Создание</span><span class="sxs-lookup"><span data-stu-id="70730-119">Compose</span></span><br><span data-ttu-id="70730-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="70730-120">Read</span></span> | [<span data-ttu-id="70730-121">Context</span><span class="sxs-lookup"><span data-stu-id="70730-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="70730-122">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="70730-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="70730-123">Enumerations</span></span>

| <span data-ttu-id="70730-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="70730-124">Enumeration</span></span> | <span data-ttu-id="70730-125">Способов</span><span class="sxs-lookup"><span data-stu-id="70730-125">Modes</span></span> | <span data-ttu-id="70730-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="70730-126">Return type</span></span> | <span data-ttu-id="70730-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="70730-127">Minimum</span></span><br><span data-ttu-id="70730-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="70730-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="70730-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="70730-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="70730-130">Создание</span><span class="sxs-lookup"><span data-stu-id="70730-130">Compose</span></span><br><span data-ttu-id="70730-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="70730-131">Read</span></span> | <span data-ttu-id="70730-132">String</span><span class="sxs-lookup"><span data-stu-id="70730-132">String</span></span> | [<span data-ttu-id="70730-133">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="70730-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="70730-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="70730-135">Создание</span><span class="sxs-lookup"><span data-stu-id="70730-135">Compose</span></span><br><span data-ttu-id="70730-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="70730-136">Read</span></span> | <span data-ttu-id="70730-137">String</span><span class="sxs-lookup"><span data-stu-id="70730-137">String</span></span> | [<span data-ttu-id="70730-138">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="70730-139">EventType</span><span class="sxs-lookup"><span data-stu-id="70730-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="70730-140">Создание</span><span class="sxs-lookup"><span data-stu-id="70730-140">Compose</span></span><br><span data-ttu-id="70730-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="70730-141">Read</span></span> | <span data-ttu-id="70730-142">String</span><span class="sxs-lookup"><span data-stu-id="70730-142">String</span></span> | [<span data-ttu-id="70730-143">1,5</span><span class="sxs-lookup"><span data-stu-id="70730-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="70730-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="70730-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="70730-145">Создание</span><span class="sxs-lookup"><span data-stu-id="70730-145">Compose</span></span><br><span data-ttu-id="70730-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="70730-146">Read</span></span> | <span data-ttu-id="70730-147">String</span><span class="sxs-lookup"><span data-stu-id="70730-147">String</span></span> | [<span data-ttu-id="70730-148">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="70730-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="70730-149">Namespaces</span></span>

<span data-ttu-id="70730-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="70730-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="70730-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="70730-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="70730-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="70730-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="70730-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="70730-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="70730-154">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-154">Type</span></span>

*   <span data-ttu-id="70730-155">String</span><span class="sxs-lookup"><span data-stu-id="70730-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70730-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="70730-156">Properties:</span></span>

|<span data-ttu-id="70730-157">Имя</span><span class="sxs-lookup"><span data-stu-id="70730-157">Name</span></span>| <span data-ttu-id="70730-158">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-158">Type</span></span>| <span data-ttu-id="70730-159">Описание</span><span class="sxs-lookup"><span data-stu-id="70730-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="70730-160">String</span><span class="sxs-lookup"><span data-stu-id="70730-160">String</span></span>|<span data-ttu-id="70730-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="70730-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="70730-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="70730-162">String</span></span>|<span data-ttu-id="70730-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="70730-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70730-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="70730-164">Requirements</span></span>

|<span data-ttu-id="70730-165">Требование</span><span class="sxs-lookup"><span data-stu-id="70730-165">Requirement</span></span>| <span data-ttu-id="70730-166">Значение</span><span class="sxs-lookup"><span data-stu-id="70730-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="70730-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="70730-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="70730-168">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-168">1.1</span></span>|
|[<span data-ttu-id="70730-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="70730-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="70730-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="70730-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="70730-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="70730-171">CoercionType: String</span></span>

<span data-ttu-id="70730-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="70730-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="70730-173">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-173">Type</span></span>

*   <span data-ttu-id="70730-174">String</span><span class="sxs-lookup"><span data-stu-id="70730-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70730-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="70730-175">Properties:</span></span>

|<span data-ttu-id="70730-176">Имя</span><span class="sxs-lookup"><span data-stu-id="70730-176">Name</span></span>| <span data-ttu-id="70730-177">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-177">Type</span></span>| <span data-ttu-id="70730-178">Описание</span><span class="sxs-lookup"><span data-stu-id="70730-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="70730-179">String</span><span class="sxs-lookup"><span data-stu-id="70730-179">String</span></span>|<span data-ttu-id="70730-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="70730-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="70730-181">String</span><span class="sxs-lookup"><span data-stu-id="70730-181">String</span></span>|<span data-ttu-id="70730-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="70730-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70730-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="70730-183">Requirements</span></span>

|<span data-ttu-id="70730-184">Требование</span><span class="sxs-lookup"><span data-stu-id="70730-184">Requirement</span></span>| <span data-ttu-id="70730-185">Значение</span><span class="sxs-lookup"><span data-stu-id="70730-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="70730-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="70730-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="70730-187">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-187">1.1</span></span>|
|[<span data-ttu-id="70730-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="70730-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="70730-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="70730-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="70730-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="70730-190">EventType: String</span></span>

<span data-ttu-id="70730-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="70730-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="70730-192">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-192">Type</span></span>

*   <span data-ttu-id="70730-193">String</span><span class="sxs-lookup"><span data-stu-id="70730-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70730-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="70730-194">Properties:</span></span>

| <span data-ttu-id="70730-195">Имя</span><span class="sxs-lookup"><span data-stu-id="70730-195">Name</span></span> | <span data-ttu-id="70730-196">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-196">Type</span></span> | <span data-ttu-id="70730-197">Описание</span><span class="sxs-lookup"><span data-stu-id="70730-197">Description</span></span> | <span data-ttu-id="70730-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="70730-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="70730-199">String</span><span class="sxs-lookup"><span data-stu-id="70730-199">String</span></span> | <span data-ttu-id="70730-200">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="70730-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="70730-201">1.7</span><span class="sxs-lookup"><span data-stu-id="70730-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="70730-202">String</span><span class="sxs-lookup"><span data-stu-id="70730-202">String</span></span> | <span data-ttu-id="70730-203">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="70730-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="70730-204">1.8</span><span class="sxs-lookup"><span data-stu-id="70730-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="70730-205">String</span><span class="sxs-lookup"><span data-stu-id="70730-205">String</span></span> | <span data-ttu-id="70730-206">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="70730-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="70730-207">1.8</span><span class="sxs-lookup"><span data-stu-id="70730-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="70730-208">String</span><span class="sxs-lookup"><span data-stu-id="70730-208">String</span></span> | <span data-ttu-id="70730-209">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="70730-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="70730-210">1.5</span><span class="sxs-lookup"><span data-stu-id="70730-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="70730-211">String</span><span class="sxs-lookup"><span data-stu-id="70730-211">String</span></span> | <span data-ttu-id="70730-212">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="70730-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="70730-213">1.7</span><span class="sxs-lookup"><span data-stu-id="70730-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="70730-214">String</span><span class="sxs-lookup"><span data-stu-id="70730-214">String</span></span> | <span data-ttu-id="70730-215">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="70730-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="70730-216">1.7</span><span class="sxs-lookup"><span data-stu-id="70730-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="70730-217">Requirements</span><span class="sxs-lookup"><span data-stu-id="70730-217">Requirements</span></span>

|<span data-ttu-id="70730-218">Требование</span><span class="sxs-lookup"><span data-stu-id="70730-218">Requirement</span></span>| <span data-ttu-id="70730-219">Значение</span><span class="sxs-lookup"><span data-stu-id="70730-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="70730-220">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="70730-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="70730-221">1.5</span><span class="sxs-lookup"><span data-stu-id="70730-221">1.5</span></span> |
|[<span data-ttu-id="70730-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="70730-222">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="70730-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="70730-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="70730-224">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="70730-224">SourceProperty: String</span></span>

<span data-ttu-id="70730-225">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="70730-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="70730-226">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-226">Type</span></span>

*   <span data-ttu-id="70730-227">String</span><span class="sxs-lookup"><span data-stu-id="70730-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70730-228">Свойства:</span><span class="sxs-lookup"><span data-stu-id="70730-228">Properties:</span></span>

|<span data-ttu-id="70730-229">Имя</span><span class="sxs-lookup"><span data-stu-id="70730-229">Name</span></span>| <span data-ttu-id="70730-230">Тип</span><span class="sxs-lookup"><span data-stu-id="70730-230">Type</span></span>| <span data-ttu-id="70730-231">Описание</span><span class="sxs-lookup"><span data-stu-id="70730-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="70730-232">String</span><span class="sxs-lookup"><span data-stu-id="70730-232">String</span></span>|<span data-ttu-id="70730-233">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="70730-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="70730-234">String</span><span class="sxs-lookup"><span data-stu-id="70730-234">String</span></span>|<span data-ttu-id="70730-235">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="70730-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70730-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="70730-236">Requirements</span></span>

|<span data-ttu-id="70730-237">Требование</span><span class="sxs-lookup"><span data-stu-id="70730-237">Requirement</span></span>| <span data-ttu-id="70730-238">Значение</span><span class="sxs-lookup"><span data-stu-id="70730-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="70730-239">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="70730-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="70730-240">1.1</span><span class="sxs-lookup"><span data-stu-id="70730-240">1.1</span></span>|
|[<span data-ttu-id="70730-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="70730-241">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="70730-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="70730-242">Compose or Read</span></span>|
