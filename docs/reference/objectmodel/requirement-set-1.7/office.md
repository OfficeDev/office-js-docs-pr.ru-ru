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
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="5ad13-103">Office (набор требований для почтового ящика 1,7)</span><span class="sxs-lookup"><span data-stu-id="5ad13-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="5ad13-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5ad13-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ad13-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="5ad13-106">Requirements</span></span>

|<span data-ttu-id="5ad13-107">Требование</span><span class="sxs-lookup"><span data-stu-id="5ad13-107">Requirement</span></span>| <span data-ttu-id="5ad13-108">Значение</span><span class="sxs-lookup"><span data-stu-id="5ad13-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ad13-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ad13-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5ad13-110">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-110">1.1</span></span>|
|[<span data-ttu-id="5ad13-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ad13-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5ad13-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="5ad13-113">Properties</span><span class="sxs-lookup"><span data-stu-id="5ad13-113">Properties</span></span>

| <span data-ttu-id="5ad13-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="5ad13-114">Property</span></span> | <span data-ttu-id="5ad13-115">Способов</span><span class="sxs-lookup"><span data-stu-id="5ad13-115">Modes</span></span> | <span data-ttu-id="5ad13-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="5ad13-116">Return type</span></span> | <span data-ttu-id="5ad13-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="5ad13-117">Minimum</span></span><br><span data-ttu-id="5ad13-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="5ad13-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5ad13-119">контекст</span><span class="sxs-lookup"><span data-stu-id="5ad13-119">context</span></span>](office.context.md) | <span data-ttu-id="5ad13-120">Создание</span><span class="sxs-lookup"><span data-stu-id="5ad13-120">Compose</span></span><br><span data-ttu-id="5ad13-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-121">Read</span></span> | [<span data-ttu-id="5ad13-122">Context</span><span class="sxs-lookup"><span data-stu-id="5ad13-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="5ad13-123">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="5ad13-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="5ad13-124">Enumerations</span></span>

| <span data-ttu-id="5ad13-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="5ad13-125">Enumeration</span></span> | <span data-ttu-id="5ad13-126">Способов</span><span class="sxs-lookup"><span data-stu-id="5ad13-126">Modes</span></span> | <span data-ttu-id="5ad13-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="5ad13-127">Return type</span></span> | <span data-ttu-id="5ad13-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="5ad13-128">Minimum</span></span><br><span data-ttu-id="5ad13-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="5ad13-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5ad13-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5ad13-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5ad13-131">Создание</span><span class="sxs-lookup"><span data-stu-id="5ad13-131">Compose</span></span><br><span data-ttu-id="5ad13-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-132">Read</span></span> | <span data-ttu-id="5ad13-133">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-133">String</span></span> | [<span data-ttu-id="5ad13-134">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5ad13-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5ad13-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5ad13-136">Создание</span><span class="sxs-lookup"><span data-stu-id="5ad13-136">Compose</span></span><br><span data-ttu-id="5ad13-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-137">Read</span></span> | <span data-ttu-id="5ad13-138">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-138">String</span></span> | [<span data-ttu-id="5ad13-139">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5ad13-140">EventType</span><span class="sxs-lookup"><span data-stu-id="5ad13-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5ad13-141">Создание</span><span class="sxs-lookup"><span data-stu-id="5ad13-141">Compose</span></span><br><span data-ttu-id="5ad13-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-142">Read</span></span> | <span data-ttu-id="5ad13-143">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-143">String</span></span> | [<span data-ttu-id="5ad13-144">1,5</span><span class="sxs-lookup"><span data-stu-id="5ad13-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="5ad13-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5ad13-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5ad13-146">Создание</span><span class="sxs-lookup"><span data-stu-id="5ad13-146">Compose</span></span><br><span data-ttu-id="5ad13-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-147">Read</span></span> | <span data-ttu-id="5ad13-148">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-148">String</span></span> | [<span data-ttu-id="5ad13-149">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="5ad13-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="5ad13-150">Namespaces</span></span>

<span data-ttu-id="5ad13-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="5ad13-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="5ad13-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="5ad13-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="5ad13-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="5ad13-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="5ad13-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ad13-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5ad13-155">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-155">Type</span></span>

*   <span data-ttu-id="5ad13-156">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5ad13-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="5ad13-157">Properties:</span></span>

|<span data-ttu-id="5ad13-158">Имя</span><span class="sxs-lookup"><span data-stu-id="5ad13-158">Name</span></span>| <span data-ttu-id="5ad13-159">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-159">Type</span></span>| <span data-ttu-id="5ad13-160">Описание</span><span class="sxs-lookup"><span data-stu-id="5ad13-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5ad13-161">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-161">String</span></span>|<span data-ttu-id="5ad13-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="5ad13-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5ad13-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="5ad13-163">String</span></span>|<span data-ttu-id="5ad13-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="5ad13-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ad13-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="5ad13-165">Requirements</span></span>

|<span data-ttu-id="5ad13-166">Требование</span><span class="sxs-lookup"><span data-stu-id="5ad13-166">Requirement</span></span>| <span data-ttu-id="5ad13-167">Значение</span><span class="sxs-lookup"><span data-stu-id="5ad13-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ad13-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ad13-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5ad13-169">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-169">1.1</span></span>|
|[<span data-ttu-id="5ad13-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ad13-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5ad13-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="5ad13-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="5ad13-172">CoercionType: String</span></span>

<span data-ttu-id="5ad13-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="5ad13-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5ad13-174">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-174">Type</span></span>

*   <span data-ttu-id="5ad13-175">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5ad13-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="5ad13-176">Properties:</span></span>

|<span data-ttu-id="5ad13-177">Имя</span><span class="sxs-lookup"><span data-stu-id="5ad13-177">Name</span></span>| <span data-ttu-id="5ad13-178">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-178">Type</span></span>| <span data-ttu-id="5ad13-179">Описание</span><span class="sxs-lookup"><span data-stu-id="5ad13-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5ad13-180">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-180">String</span></span>|<span data-ttu-id="5ad13-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="5ad13-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5ad13-182">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-182">String</span></span>|<span data-ttu-id="5ad13-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="5ad13-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ad13-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="5ad13-184">Requirements</span></span>

|<span data-ttu-id="5ad13-185">Требование</span><span class="sxs-lookup"><span data-stu-id="5ad13-185">Requirement</span></span>| <span data-ttu-id="5ad13-186">Значение</span><span class="sxs-lookup"><span data-stu-id="5ad13-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ad13-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ad13-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5ad13-188">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-188">1.1</span></span>|
|[<span data-ttu-id="5ad13-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ad13-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5ad13-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="5ad13-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="5ad13-191">EventType: String</span></span>

<span data-ttu-id="5ad13-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="5ad13-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5ad13-193">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-193">Type</span></span>

*   <span data-ttu-id="5ad13-194">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5ad13-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="5ad13-195">Properties:</span></span>

| <span data-ttu-id="5ad13-196">Имя</span><span class="sxs-lookup"><span data-stu-id="5ad13-196">Name</span></span> | <span data-ttu-id="5ad13-197">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-197">Type</span></span> | <span data-ttu-id="5ad13-198">Описание</span><span class="sxs-lookup"><span data-stu-id="5ad13-198">Description</span></span> | <span data-ttu-id="5ad13-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="5ad13-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="5ad13-200">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-200">String</span></span> | <span data-ttu-id="5ad13-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="5ad13-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5ad13-202">1.7</span><span class="sxs-lookup"><span data-stu-id="5ad13-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="5ad13-203">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-203">String</span></span> | <span data-ttu-id="5ad13-204">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="5ad13-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5ad13-205">1.5</span><span class="sxs-lookup"><span data-stu-id="5ad13-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5ad13-206">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-206">String</span></span> | <span data-ttu-id="5ad13-207">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="5ad13-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5ad13-208">1.7</span><span class="sxs-lookup"><span data-stu-id="5ad13-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5ad13-209">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-209">String</span></span> | <span data-ttu-id="5ad13-210">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="5ad13-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5ad13-211">1.7</span><span class="sxs-lookup"><span data-stu-id="5ad13-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ad13-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="5ad13-212">Requirements</span></span>

|<span data-ttu-id="5ad13-213">Требование</span><span class="sxs-lookup"><span data-stu-id="5ad13-213">Requirement</span></span>| <span data-ttu-id="5ad13-214">Значение</span><span class="sxs-lookup"><span data-stu-id="5ad13-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ad13-215">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ad13-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5ad13-216">1.5</span><span class="sxs-lookup"><span data-stu-id="5ad13-216">1.5</span></span> |
|[<span data-ttu-id="5ad13-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ad13-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5ad13-218">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="5ad13-219">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="5ad13-219">SourceProperty: String</span></span>

<span data-ttu-id="5ad13-220">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="5ad13-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5ad13-221">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-221">Type</span></span>

*   <span data-ttu-id="5ad13-222">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5ad13-223">Свойства:</span><span class="sxs-lookup"><span data-stu-id="5ad13-223">Properties:</span></span>

|<span data-ttu-id="5ad13-224">Имя</span><span class="sxs-lookup"><span data-stu-id="5ad13-224">Name</span></span>| <span data-ttu-id="5ad13-225">Тип</span><span class="sxs-lookup"><span data-stu-id="5ad13-225">Type</span></span>| <span data-ttu-id="5ad13-226">Описание</span><span class="sxs-lookup"><span data-stu-id="5ad13-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5ad13-227">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-227">String</span></span>|<span data-ttu-id="5ad13-228">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ad13-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5ad13-229">String</span><span class="sxs-lookup"><span data-stu-id="5ad13-229">String</span></span>|<span data-ttu-id="5ad13-230">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ad13-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ad13-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="5ad13-231">Requirements</span></span>

|<span data-ttu-id="5ad13-232">Требование</span><span class="sxs-lookup"><span data-stu-id="5ad13-232">Requirement</span></span>| <span data-ttu-id="5ad13-233">Значение</span><span class="sxs-lookup"><span data-stu-id="5ad13-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ad13-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ad13-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5ad13-235">1.1</span><span class="sxs-lookup"><span data-stu-id="5ad13-235">1.1</span></span>|
|[<span data-ttu-id="5ad13-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ad13-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5ad13-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ad13-237">Compose or Read</span></span>|
