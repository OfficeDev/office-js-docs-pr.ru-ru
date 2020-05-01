---
title: Пространство имен Office — набор обязательных элементов 1,7
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 7991fd56097bbdebbfd4d4494a900626a1d3e02b
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891252"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="646ba-103">Office (набор требований для почтового ящика 1,7)</span><span class="sxs-lookup"><span data-stu-id="646ba-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="646ba-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="646ba-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="646ba-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="646ba-106">Requirements</span></span>

|<span data-ttu-id="646ba-107">Требование</span><span class="sxs-lookup"><span data-stu-id="646ba-107">Requirement</span></span>| <span data-ttu-id="646ba-108">Значение</span><span class="sxs-lookup"><span data-stu-id="646ba-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="646ba-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="646ba-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="646ba-110">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-110">1.1</span></span>|
|[<span data-ttu-id="646ba-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="646ba-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="646ba-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="646ba-113">Properties</span><span class="sxs-lookup"><span data-stu-id="646ba-113">Properties</span></span>

| <span data-ttu-id="646ba-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="646ba-114">Property</span></span> | <span data-ttu-id="646ba-115">Способов</span><span class="sxs-lookup"><span data-stu-id="646ba-115">Modes</span></span> | <span data-ttu-id="646ba-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="646ba-116">Return type</span></span> | <span data-ttu-id="646ba-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="646ba-117">Minimum</span></span><br><span data-ttu-id="646ba-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="646ba-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="646ba-119">контекст</span><span class="sxs-lookup"><span data-stu-id="646ba-119">context</span></span>](office.context.md) | <span data-ttu-id="646ba-120">Создание</span><span class="sxs-lookup"><span data-stu-id="646ba-120">Compose</span></span><br><span data-ttu-id="646ba-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-121">Read</span></span> | [<span data-ttu-id="646ba-122">Context</span><span class="sxs-lookup"><span data-stu-id="646ba-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="646ba-123">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="646ba-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="646ba-124">Enumerations</span></span>

| <span data-ttu-id="646ba-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="646ba-125">Enumeration</span></span> | <span data-ttu-id="646ba-126">Способов</span><span class="sxs-lookup"><span data-stu-id="646ba-126">Modes</span></span> | <span data-ttu-id="646ba-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="646ba-127">Return type</span></span> | <span data-ttu-id="646ba-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="646ba-128">Minimum</span></span><br><span data-ttu-id="646ba-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="646ba-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="646ba-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="646ba-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="646ba-131">Создание</span><span class="sxs-lookup"><span data-stu-id="646ba-131">Compose</span></span><br><span data-ttu-id="646ba-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-132">Read</span></span> | <span data-ttu-id="646ba-133">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-133">String</span></span> | [<span data-ttu-id="646ba-134">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="646ba-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="646ba-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="646ba-136">Создание</span><span class="sxs-lookup"><span data-stu-id="646ba-136">Compose</span></span><br><span data-ttu-id="646ba-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-137">Read</span></span> | <span data-ttu-id="646ba-138">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-138">String</span></span> | [<span data-ttu-id="646ba-139">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="646ba-140">EventType</span><span class="sxs-lookup"><span data-stu-id="646ba-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="646ba-141">Создание</span><span class="sxs-lookup"><span data-stu-id="646ba-141">Compose</span></span><br><span data-ttu-id="646ba-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-142">Read</span></span> | <span data-ttu-id="646ba-143">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-143">String</span></span> | [<span data-ttu-id="646ba-144">1,5</span><span class="sxs-lookup"><span data-stu-id="646ba-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="646ba-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="646ba-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="646ba-146">Создание</span><span class="sxs-lookup"><span data-stu-id="646ba-146">Compose</span></span><br><span data-ttu-id="646ba-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-147">Read</span></span> | <span data-ttu-id="646ba-148">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-148">String</span></span> | [<span data-ttu-id="646ba-149">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="646ba-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="646ba-150">Namespaces</span></span>

<span data-ttu-id="646ba-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="646ba-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="646ba-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="646ba-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="646ba-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="646ba-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="646ba-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="646ba-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="646ba-155">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-155">Type</span></span>

*   <span data-ttu-id="646ba-156">String</span><span class="sxs-lookup"><span data-stu-id="646ba-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="646ba-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="646ba-157">Properties:</span></span>

|<span data-ttu-id="646ba-158">Имя</span><span class="sxs-lookup"><span data-stu-id="646ba-158">Name</span></span>| <span data-ttu-id="646ba-159">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-159">Type</span></span>| <span data-ttu-id="646ba-160">Описание</span><span class="sxs-lookup"><span data-stu-id="646ba-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="646ba-161">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-161">String</span></span>|<span data-ttu-id="646ba-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="646ba-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="646ba-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="646ba-163">String</span></span>|<span data-ttu-id="646ba-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="646ba-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="646ba-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="646ba-165">Requirements</span></span>

|<span data-ttu-id="646ba-166">Требование</span><span class="sxs-lookup"><span data-stu-id="646ba-166">Requirement</span></span>| <span data-ttu-id="646ba-167">Значение</span><span class="sxs-lookup"><span data-stu-id="646ba-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="646ba-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="646ba-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="646ba-169">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-169">1.1</span></span>|
|[<span data-ttu-id="646ba-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="646ba-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="646ba-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="646ba-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="646ba-172">CoercionType: String</span></span>

<span data-ttu-id="646ba-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="646ba-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="646ba-174">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-174">Type</span></span>

*   <span data-ttu-id="646ba-175">String</span><span class="sxs-lookup"><span data-stu-id="646ba-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="646ba-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="646ba-176">Properties:</span></span>

|<span data-ttu-id="646ba-177">Имя</span><span class="sxs-lookup"><span data-stu-id="646ba-177">Name</span></span>| <span data-ttu-id="646ba-178">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-178">Type</span></span>| <span data-ttu-id="646ba-179">Описание</span><span class="sxs-lookup"><span data-stu-id="646ba-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="646ba-180">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-180">String</span></span>|<span data-ttu-id="646ba-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="646ba-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="646ba-182">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-182">String</span></span>|<span data-ttu-id="646ba-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="646ba-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="646ba-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="646ba-184">Requirements</span></span>

|<span data-ttu-id="646ba-185">Требование</span><span class="sxs-lookup"><span data-stu-id="646ba-185">Requirement</span></span>| <span data-ttu-id="646ba-186">Значение</span><span class="sxs-lookup"><span data-stu-id="646ba-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="646ba-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="646ba-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="646ba-188">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-188">1.1</span></span>|
|[<span data-ttu-id="646ba-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="646ba-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="646ba-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="646ba-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="646ba-191">EventType: String</span></span>

<span data-ttu-id="646ba-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="646ba-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="646ba-193">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-193">Type</span></span>

*   <span data-ttu-id="646ba-194">String</span><span class="sxs-lookup"><span data-stu-id="646ba-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="646ba-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="646ba-195">Properties:</span></span>

| <span data-ttu-id="646ba-196">Имя</span><span class="sxs-lookup"><span data-stu-id="646ba-196">Name</span></span> | <span data-ttu-id="646ba-197">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-197">Type</span></span> | <span data-ttu-id="646ba-198">Описание</span><span class="sxs-lookup"><span data-stu-id="646ba-198">Description</span></span> | <span data-ttu-id="646ba-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="646ba-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="646ba-200">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-200">String</span></span> | <span data-ttu-id="646ba-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="646ba-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="646ba-202">1.7</span><span class="sxs-lookup"><span data-stu-id="646ba-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="646ba-203">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-203">String</span></span> | <span data-ttu-id="646ba-204">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="646ba-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="646ba-205">1.5</span><span class="sxs-lookup"><span data-stu-id="646ba-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="646ba-206">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-206">String</span></span> | <span data-ttu-id="646ba-207">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="646ba-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="646ba-208">1.7</span><span class="sxs-lookup"><span data-stu-id="646ba-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="646ba-209">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-209">String</span></span> | <span data-ttu-id="646ba-210">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="646ba-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="646ba-211">1.7</span><span class="sxs-lookup"><span data-stu-id="646ba-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="646ba-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="646ba-212">Requirements</span></span>

|<span data-ttu-id="646ba-213">Требование</span><span class="sxs-lookup"><span data-stu-id="646ba-213">Requirement</span></span>| <span data-ttu-id="646ba-214">Значение</span><span class="sxs-lookup"><span data-stu-id="646ba-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="646ba-215">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="646ba-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="646ba-216">1.5</span><span class="sxs-lookup"><span data-stu-id="646ba-216">1.5</span></span> |
|[<span data-ttu-id="646ba-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="646ba-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="646ba-218">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="646ba-219">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="646ba-219">SourceProperty: String</span></span>

<span data-ttu-id="646ba-220">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="646ba-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="646ba-221">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-221">Type</span></span>

*   <span data-ttu-id="646ba-222">String</span><span class="sxs-lookup"><span data-stu-id="646ba-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="646ba-223">Свойства:</span><span class="sxs-lookup"><span data-stu-id="646ba-223">Properties:</span></span>

|<span data-ttu-id="646ba-224">Имя</span><span class="sxs-lookup"><span data-stu-id="646ba-224">Name</span></span>| <span data-ttu-id="646ba-225">Тип</span><span class="sxs-lookup"><span data-stu-id="646ba-225">Type</span></span>| <span data-ttu-id="646ba-226">Описание</span><span class="sxs-lookup"><span data-stu-id="646ba-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="646ba-227">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-227">String</span></span>|<span data-ttu-id="646ba-228">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="646ba-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="646ba-229">Строка</span><span class="sxs-lookup"><span data-stu-id="646ba-229">String</span></span>|<span data-ttu-id="646ba-230">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="646ba-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="646ba-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="646ba-231">Requirements</span></span>

|<span data-ttu-id="646ba-232">Требование</span><span class="sxs-lookup"><span data-stu-id="646ba-232">Requirement</span></span>| <span data-ttu-id="646ba-233">Значение</span><span class="sxs-lookup"><span data-stu-id="646ba-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="646ba-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="646ba-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="646ba-235">1.1</span><span class="sxs-lookup"><span data-stu-id="646ba-235">1.1</span></span>|
|[<span data-ttu-id="646ba-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="646ba-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="646ba-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="646ba-237">Compose or Read</span></span>|
