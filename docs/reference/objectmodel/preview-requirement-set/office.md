---
title: Пространство имен Office — Предварительная версия набора требований
description: Объектная модель для пространства имен верхнего уровня API надстроек Outlook (Предварительная версия API почтовых ящиков).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 40623c02fae820926d9162903320f30e5a424544
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720276"
---
# <a name="office"></a><span data-ttu-id="16cce-103">Office</span><span class="sxs-lookup"><span data-stu-id="16cce-103">Office</span></span>

<span data-ttu-id="16cce-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="16cce-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="16cce-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="16cce-106">Requirements</span></span>

|<span data-ttu-id="16cce-107">Требование</span><span class="sxs-lookup"><span data-stu-id="16cce-107">Requirement</span></span>| <span data-ttu-id="16cce-108">Значение</span><span class="sxs-lookup"><span data-stu-id="16cce-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="16cce-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="16cce-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16cce-110">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-110">1.1</span></span>|
|[<span data-ttu-id="16cce-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="16cce-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16cce-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="16cce-113">Properties</span><span class="sxs-lookup"><span data-stu-id="16cce-113">Properties</span></span>

| <span data-ttu-id="16cce-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="16cce-114">Property</span></span> | <span data-ttu-id="16cce-115">Способов</span><span class="sxs-lookup"><span data-stu-id="16cce-115">Modes</span></span> | <span data-ttu-id="16cce-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="16cce-116">Return type</span></span> | <span data-ttu-id="16cce-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="16cce-117">Minimum</span></span><br><span data-ttu-id="16cce-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="16cce-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="16cce-119">контекст</span><span class="sxs-lookup"><span data-stu-id="16cce-119">context</span></span>](office.context.md) | <span data-ttu-id="16cce-120">Создание</span><span class="sxs-lookup"><span data-stu-id="16cce-120">Compose</span></span><br><span data-ttu-id="16cce-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-121">Read</span></span> | [<span data-ttu-id="16cce-122">Context</span><span class="sxs-lookup"><span data-stu-id="16cce-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="16cce-123">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="16cce-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="16cce-124">Enumerations</span></span>

| <span data-ttu-id="16cce-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="16cce-125">Enumeration</span></span> | <span data-ttu-id="16cce-126">Способов</span><span class="sxs-lookup"><span data-stu-id="16cce-126">Modes</span></span> | <span data-ttu-id="16cce-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="16cce-127">Return type</span></span> | <span data-ttu-id="16cce-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="16cce-128">Minimum</span></span><br><span data-ttu-id="16cce-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="16cce-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="16cce-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="16cce-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="16cce-131">Создание</span><span class="sxs-lookup"><span data-stu-id="16cce-131">Compose</span></span><br><span data-ttu-id="16cce-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-132">Read</span></span> | <span data-ttu-id="16cce-133">String</span><span class="sxs-lookup"><span data-stu-id="16cce-133">String</span></span> | [<span data-ttu-id="16cce-134">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16cce-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="16cce-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="16cce-136">Создание</span><span class="sxs-lookup"><span data-stu-id="16cce-136">Compose</span></span><br><span data-ttu-id="16cce-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-137">Read</span></span> | <span data-ttu-id="16cce-138">String</span><span class="sxs-lookup"><span data-stu-id="16cce-138">String</span></span> | [<span data-ttu-id="16cce-139">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16cce-140">EventType</span><span class="sxs-lookup"><span data-stu-id="16cce-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="16cce-141">Создание</span><span class="sxs-lookup"><span data-stu-id="16cce-141">Compose</span></span><br><span data-ttu-id="16cce-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-142">Read</span></span> | <span data-ttu-id="16cce-143">String</span><span class="sxs-lookup"><span data-stu-id="16cce-143">String</span></span> | [<span data-ttu-id="16cce-144">1,5</span><span class="sxs-lookup"><span data-stu-id="16cce-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="16cce-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="16cce-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="16cce-146">Создание</span><span class="sxs-lookup"><span data-stu-id="16cce-146">Compose</span></span><br><span data-ttu-id="16cce-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-147">Read</span></span> | <span data-ttu-id="16cce-148">String</span><span class="sxs-lookup"><span data-stu-id="16cce-148">String</span></span> | [<span data-ttu-id="16cce-149">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="16cce-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="16cce-150">Namespaces</span></span>

<span data-ttu-id="16cce-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="16cce-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="16cce-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="16cce-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="16cce-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="16cce-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="16cce-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="16cce-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="16cce-155">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-155">Type</span></span>

*   <span data-ttu-id="16cce-156">String</span><span class="sxs-lookup"><span data-stu-id="16cce-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16cce-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="16cce-157">Properties:</span></span>

|<span data-ttu-id="16cce-158">Имя</span><span class="sxs-lookup"><span data-stu-id="16cce-158">Name</span></span>| <span data-ttu-id="16cce-159">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-159">Type</span></span>| <span data-ttu-id="16cce-160">Описание</span><span class="sxs-lookup"><span data-stu-id="16cce-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="16cce-161">String</span><span class="sxs-lookup"><span data-stu-id="16cce-161">String</span></span>|<span data-ttu-id="16cce-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="16cce-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="16cce-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="16cce-163">String</span></span>|<span data-ttu-id="16cce-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="16cce-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16cce-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="16cce-165">Requirements</span></span>

|<span data-ttu-id="16cce-166">Требование</span><span class="sxs-lookup"><span data-stu-id="16cce-166">Requirement</span></span>| <span data-ttu-id="16cce-167">Значение</span><span class="sxs-lookup"><span data-stu-id="16cce-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="16cce-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="16cce-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16cce-169">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-169">1.1</span></span>|
|[<span data-ttu-id="16cce-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="16cce-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16cce-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="16cce-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="16cce-172">CoercionType: String</span></span>

<span data-ttu-id="16cce-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="16cce-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="16cce-174">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-174">Type</span></span>

*   <span data-ttu-id="16cce-175">String</span><span class="sxs-lookup"><span data-stu-id="16cce-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16cce-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="16cce-176">Properties:</span></span>

|<span data-ttu-id="16cce-177">Имя</span><span class="sxs-lookup"><span data-stu-id="16cce-177">Name</span></span>| <span data-ttu-id="16cce-178">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-178">Type</span></span>| <span data-ttu-id="16cce-179">Описание</span><span class="sxs-lookup"><span data-stu-id="16cce-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="16cce-180">String</span><span class="sxs-lookup"><span data-stu-id="16cce-180">String</span></span>|<span data-ttu-id="16cce-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="16cce-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="16cce-182">String</span><span class="sxs-lookup"><span data-stu-id="16cce-182">String</span></span>|<span data-ttu-id="16cce-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="16cce-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16cce-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="16cce-184">Requirements</span></span>

|<span data-ttu-id="16cce-185">Требование</span><span class="sxs-lookup"><span data-stu-id="16cce-185">Requirement</span></span>| <span data-ttu-id="16cce-186">Значение</span><span class="sxs-lookup"><span data-stu-id="16cce-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="16cce-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="16cce-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16cce-188">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-188">1.1</span></span>|
|[<span data-ttu-id="16cce-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="16cce-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16cce-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="16cce-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="16cce-191">EventType: String</span></span>

<span data-ttu-id="16cce-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="16cce-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="16cce-193">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-193">Type</span></span>

*   <span data-ttu-id="16cce-194">String</span><span class="sxs-lookup"><span data-stu-id="16cce-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16cce-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="16cce-195">Properties:</span></span>

| <span data-ttu-id="16cce-196">Имя</span><span class="sxs-lookup"><span data-stu-id="16cce-196">Name</span></span> | <span data-ttu-id="16cce-197">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-197">Type</span></span> | <span data-ttu-id="16cce-198">Описание</span><span class="sxs-lookup"><span data-stu-id="16cce-198">Description</span></span> | <span data-ttu-id="16cce-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="16cce-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="16cce-200">String</span><span class="sxs-lookup"><span data-stu-id="16cce-200">String</span></span> | <span data-ttu-id="16cce-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="16cce-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="16cce-202">1.7</span><span class="sxs-lookup"><span data-stu-id="16cce-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="16cce-203">String</span><span class="sxs-lookup"><span data-stu-id="16cce-203">String</span></span> | <span data-ttu-id="16cce-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="16cce-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="16cce-205">1.8</span><span class="sxs-lookup"><span data-stu-id="16cce-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="16cce-206">String</span><span class="sxs-lookup"><span data-stu-id="16cce-206">String</span></span> | <span data-ttu-id="16cce-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="16cce-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="16cce-208">1.8</span><span class="sxs-lookup"><span data-stu-id="16cce-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="16cce-209">String</span><span class="sxs-lookup"><span data-stu-id="16cce-209">String</span></span> | <span data-ttu-id="16cce-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="16cce-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="16cce-211">1.5</span><span class="sxs-lookup"><span data-stu-id="16cce-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="16cce-212">String</span><span class="sxs-lookup"><span data-stu-id="16cce-212">String</span></span> | <span data-ttu-id="16cce-213">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="16cce-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="16cce-214">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="16cce-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="16cce-215">String</span><span class="sxs-lookup"><span data-stu-id="16cce-215">String</span></span> | <span data-ttu-id="16cce-216">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="16cce-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="16cce-217">1.7</span><span class="sxs-lookup"><span data-stu-id="16cce-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="16cce-218">String</span><span class="sxs-lookup"><span data-stu-id="16cce-218">String</span></span> | <span data-ttu-id="16cce-219">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="16cce-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="16cce-220">1.7</span><span class="sxs-lookup"><span data-stu-id="16cce-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="16cce-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="16cce-221">Requirements</span></span>

|<span data-ttu-id="16cce-222">Требование</span><span class="sxs-lookup"><span data-stu-id="16cce-222">Requirement</span></span>| <span data-ttu-id="16cce-223">Значение</span><span class="sxs-lookup"><span data-stu-id="16cce-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="16cce-224">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="16cce-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16cce-225">1.5</span><span class="sxs-lookup"><span data-stu-id="16cce-225">1.5</span></span> |
|[<span data-ttu-id="16cce-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="16cce-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16cce-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="16cce-228">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="16cce-228">SourceProperty: String</span></span>

<span data-ttu-id="16cce-229">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="16cce-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="16cce-230">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-230">Type</span></span>

*   <span data-ttu-id="16cce-231">String</span><span class="sxs-lookup"><span data-stu-id="16cce-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16cce-232">Свойства:</span><span class="sxs-lookup"><span data-stu-id="16cce-232">Properties:</span></span>

|<span data-ttu-id="16cce-233">Имя</span><span class="sxs-lookup"><span data-stu-id="16cce-233">Name</span></span>| <span data-ttu-id="16cce-234">Тип</span><span class="sxs-lookup"><span data-stu-id="16cce-234">Type</span></span>| <span data-ttu-id="16cce-235">Описание</span><span class="sxs-lookup"><span data-stu-id="16cce-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="16cce-236">String</span><span class="sxs-lookup"><span data-stu-id="16cce-236">String</span></span>|<span data-ttu-id="16cce-237">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="16cce-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="16cce-238">String</span><span class="sxs-lookup"><span data-stu-id="16cce-238">String</span></span>|<span data-ttu-id="16cce-239">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="16cce-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16cce-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="16cce-240">Requirements</span></span>

|<span data-ttu-id="16cce-241">Требование</span><span class="sxs-lookup"><span data-stu-id="16cce-241">Requirement</span></span>| <span data-ttu-id="16cce-242">Значение</span><span class="sxs-lookup"><span data-stu-id="16cce-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="16cce-243">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="16cce-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16cce-244">1.1</span><span class="sxs-lookup"><span data-stu-id="16cce-244">1.1</span></span>|
|[<span data-ttu-id="16cce-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="16cce-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16cce-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="16cce-246">Compose or Read</span></span>|
