---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9bfff9c45cb157d2dcd42997a01f5ada40aecfa0
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814572"
---
# <a name="office"></a><span data-ttu-id="ebed9-102">Office</span><span class="sxs-lookup"><span data-stu-id="ebed9-102">Office</span></span>

<span data-ttu-id="ebed9-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ebed9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ebed9-105">Требования</span><span class="sxs-lookup"><span data-stu-id="ebed9-105">Requirements</span></span>

|<span data-ttu-id="ebed9-106">Требование</span><span class="sxs-lookup"><span data-stu-id="ebed9-106">Requirement</span></span>| <span data-ttu-id="ebed9-107">Значение</span><span class="sxs-lookup"><span data-stu-id="ebed9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebed9-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ebed9-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebed9-109">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-109">1.1</span></span>|
|[<span data-ttu-id="ebed9-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ebed9-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ebed9-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ebed9-112">Properties</span><span class="sxs-lookup"><span data-stu-id="ebed9-112">Properties</span></span>

| <span data-ttu-id="ebed9-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="ebed9-113">Property</span></span> | <span data-ttu-id="ebed9-114">Способов</span><span class="sxs-lookup"><span data-stu-id="ebed9-114">Modes</span></span> | <span data-ttu-id="ebed9-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="ebed9-115">Return type</span></span> | <span data-ttu-id="ebed9-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="ebed9-116">Minimum</span></span><br><span data-ttu-id="ebed9-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="ebed9-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ebed9-118">контекст</span><span class="sxs-lookup"><span data-stu-id="ebed9-118">context</span></span>](office.context.md) | <span data-ttu-id="ebed9-119">Создание</span><span class="sxs-lookup"><span data-stu-id="ebed9-119">Compose</span></span><br><span data-ttu-id="ebed9-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-120">Read</span></span> | [<span data-ttu-id="ebed9-121">Context</span><span class="sxs-lookup"><span data-stu-id="ebed9-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="ebed9-122">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="ebed9-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="ebed9-123">Enumerations</span></span>

| <span data-ttu-id="ebed9-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="ebed9-124">Enumeration</span></span> | <span data-ttu-id="ebed9-125">Способов</span><span class="sxs-lookup"><span data-stu-id="ebed9-125">Modes</span></span> | <span data-ttu-id="ebed9-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="ebed9-126">Return type</span></span> | <span data-ttu-id="ebed9-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="ebed9-127">Minimum</span></span><br><span data-ttu-id="ebed9-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="ebed9-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ebed9-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ebed9-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ebed9-130">Создание</span><span class="sxs-lookup"><span data-stu-id="ebed9-130">Compose</span></span><br><span data-ttu-id="ebed9-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-131">Read</span></span> | <span data-ttu-id="ebed9-132">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-132">String</span></span> | [<span data-ttu-id="ebed9-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebed9-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ebed9-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ebed9-135">Создание</span><span class="sxs-lookup"><span data-stu-id="ebed9-135">Compose</span></span><br><span data-ttu-id="ebed9-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-136">Read</span></span> | <span data-ttu-id="ebed9-137">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-137">String</span></span> | [<span data-ttu-id="ebed9-138">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebed9-139">EventType</span><span class="sxs-lookup"><span data-stu-id="ebed9-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ebed9-140">Создание</span><span class="sxs-lookup"><span data-stu-id="ebed9-140">Compose</span></span><br><span data-ttu-id="ebed9-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-141">Read</span></span> | <span data-ttu-id="ebed9-142">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-142">String</span></span> | [<span data-ttu-id="ebed9-143">1,5</span><span class="sxs-lookup"><span data-stu-id="ebed9-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ebed9-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ebed9-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ebed9-145">Создание</span><span class="sxs-lookup"><span data-stu-id="ebed9-145">Compose</span></span><br><span data-ttu-id="ebed9-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-146">Read</span></span> | <span data-ttu-id="ebed9-147">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-147">String</span></span> | [<span data-ttu-id="ebed9-148">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="ebed9-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="ebed9-149">Namespaces</span></span>

<span data-ttu-id="ebed9-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="ebed9-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="ebed9-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="ebed9-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ebed9-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="ebed9-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="ebed9-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="ebed9-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ebed9-154">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-154">Type</span></span>

*   <span data-ttu-id="ebed9-155">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ebed9-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ebed9-156">Properties:</span></span>

|<span data-ttu-id="ebed9-157">Имя</span><span class="sxs-lookup"><span data-stu-id="ebed9-157">Name</span></span>| <span data-ttu-id="ebed9-158">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-158">Type</span></span>| <span data-ttu-id="ebed9-159">Описание</span><span class="sxs-lookup"><span data-stu-id="ebed9-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ebed9-160">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-160">String</span></span>|<span data-ttu-id="ebed9-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="ebed9-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ebed9-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="ebed9-162">String</span></span>|<span data-ttu-id="ebed9-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="ebed9-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ebed9-164">Требования</span><span class="sxs-lookup"><span data-stu-id="ebed9-164">Requirements</span></span>

|<span data-ttu-id="ebed9-165">Требование</span><span class="sxs-lookup"><span data-stu-id="ebed9-165">Requirement</span></span>| <span data-ttu-id="ebed9-166">Значение</span><span class="sxs-lookup"><span data-stu-id="ebed9-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebed9-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ebed9-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebed9-168">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-168">1.1</span></span>|
|[<span data-ttu-id="ebed9-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ebed9-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ebed9-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="ebed9-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="ebed9-171">CoercionType: String</span></span>

<span data-ttu-id="ebed9-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ebed9-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ebed9-173">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-173">Type</span></span>

*   <span data-ttu-id="ebed9-174">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ebed9-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ebed9-175">Properties:</span></span>

|<span data-ttu-id="ebed9-176">Имя</span><span class="sxs-lookup"><span data-stu-id="ebed9-176">Name</span></span>| <span data-ttu-id="ebed9-177">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-177">Type</span></span>| <span data-ttu-id="ebed9-178">Описание</span><span class="sxs-lookup"><span data-stu-id="ebed9-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ebed9-179">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-179">String</span></span>|<span data-ttu-id="ebed9-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="ebed9-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ebed9-181">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-181">String</span></span>|<span data-ttu-id="ebed9-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="ebed9-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ebed9-183">Требования</span><span class="sxs-lookup"><span data-stu-id="ebed9-183">Requirements</span></span>

|<span data-ttu-id="ebed9-184">Требование</span><span class="sxs-lookup"><span data-stu-id="ebed9-184">Requirement</span></span>| <span data-ttu-id="ebed9-185">Значение</span><span class="sxs-lookup"><span data-stu-id="ebed9-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebed9-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ebed9-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebed9-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-187">1.1</span></span>|
|[<span data-ttu-id="ebed9-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ebed9-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ebed9-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="ebed9-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="ebed9-190">EventType: String</span></span>

<span data-ttu-id="ebed9-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="ebed9-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ebed9-192">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-192">Type</span></span>

*   <span data-ttu-id="ebed9-193">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ebed9-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ebed9-194">Properties:</span></span>

| <span data-ttu-id="ebed9-195">Имя</span><span class="sxs-lookup"><span data-stu-id="ebed9-195">Name</span></span> | <span data-ttu-id="ebed9-196">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-196">Type</span></span> | <span data-ttu-id="ebed9-197">Описание</span><span class="sxs-lookup"><span data-stu-id="ebed9-197">Description</span></span> | <span data-ttu-id="ebed9-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="ebed9-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="ebed9-199">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-199">String</span></span> | <span data-ttu-id="ebed9-200">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="ebed9-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ebed9-201">1.7</span><span class="sxs-lookup"><span data-stu-id="ebed9-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="ebed9-202">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-202">String</span></span> | <span data-ttu-id="ebed9-203">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="ebed9-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ebed9-204">1.5</span><span class="sxs-lookup"><span data-stu-id="ebed9-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ebed9-205">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-205">String</span></span> | <span data-ttu-id="ebed9-206">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="ebed9-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ebed9-207">1.7</span><span class="sxs-lookup"><span data-stu-id="ebed9-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ebed9-208">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-208">String</span></span> | <span data-ttu-id="ebed9-209">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="ebed9-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ebed9-210">1.7</span><span class="sxs-lookup"><span data-stu-id="ebed9-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ebed9-211">Требования</span><span class="sxs-lookup"><span data-stu-id="ebed9-211">Requirements</span></span>

|<span data-ttu-id="ebed9-212">Требование</span><span class="sxs-lookup"><span data-stu-id="ebed9-212">Requirement</span></span>| <span data-ttu-id="ebed9-213">Значение</span><span class="sxs-lookup"><span data-stu-id="ebed9-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebed9-214">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ebed9-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebed9-215">1.5</span><span class="sxs-lookup"><span data-stu-id="ebed9-215">1.5</span></span> |
|[<span data-ttu-id="ebed9-216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ebed9-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ebed9-217">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="ebed9-218">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="ebed9-218">SourceProperty: String</span></span>

<span data-ttu-id="ebed9-219">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ebed9-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ebed9-220">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-220">Type</span></span>

*   <span data-ttu-id="ebed9-221">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ebed9-222">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ebed9-222">Properties:</span></span>

|<span data-ttu-id="ebed9-223">Имя</span><span class="sxs-lookup"><span data-stu-id="ebed9-223">Name</span></span>| <span data-ttu-id="ebed9-224">Тип</span><span class="sxs-lookup"><span data-stu-id="ebed9-224">Type</span></span>| <span data-ttu-id="ebed9-225">Описание</span><span class="sxs-lookup"><span data-stu-id="ebed9-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ebed9-226">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-226">String</span></span>|<span data-ttu-id="ebed9-227">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="ebed9-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ebed9-228">String</span><span class="sxs-lookup"><span data-stu-id="ebed9-228">String</span></span>|<span data-ttu-id="ebed9-229">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="ebed9-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ebed9-230">Требования</span><span class="sxs-lookup"><span data-stu-id="ebed9-230">Requirements</span></span>

|<span data-ttu-id="ebed9-231">Требование</span><span class="sxs-lookup"><span data-stu-id="ebed9-231">Requirement</span></span>| <span data-ttu-id="ebed9-232">Значение</span><span class="sxs-lookup"><span data-stu-id="ebed9-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebed9-233">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ebed9-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebed9-234">1.1</span><span class="sxs-lookup"><span data-stu-id="ebed9-234">1.1</span></span>|
|[<span data-ttu-id="ebed9-235">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ebed9-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ebed9-236">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ebed9-236">Compose or Read</span></span>|
