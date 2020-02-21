---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 23f3fb705c03eabd8ee7fce53f4c89a48128672f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165350"
---
# <a name="office"></a><span data-ttu-id="2a6e7-102">Office</span><span class="sxs-lookup"><span data-stu-id="2a6e7-102">Office</span></span>

<span data-ttu-id="2a6e7-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2a6e7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a6e7-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a6e7-105">Requirements</span></span>

|<span data-ttu-id="2a6e7-106">Требование</span><span class="sxs-lookup"><span data-stu-id="2a6e7-106">Requirement</span></span>| <span data-ttu-id="2a6e7-107">Значение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a6e7-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a6e7-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a6e7-109">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-109">1.1</span></span>|
|[<span data-ttu-id="2a6e7-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a6e7-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a6e7-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2a6e7-112">Properties</span><span class="sxs-lookup"><span data-stu-id="2a6e7-112">Properties</span></span>

| <span data-ttu-id="2a6e7-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="2a6e7-113">Property</span></span> | <span data-ttu-id="2a6e7-114">Способов</span><span class="sxs-lookup"><span data-stu-id="2a6e7-114">Modes</span></span> | <span data-ttu-id="2a6e7-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="2a6e7-115">Return type</span></span> | <span data-ttu-id="2a6e7-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="2a6e7-116">Minimum</span></span><br><span data-ttu-id="2a6e7-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="2a6e7-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2a6e7-118">контекст</span><span class="sxs-lookup"><span data-stu-id="2a6e7-118">context</span></span>](office.context.md) | <span data-ttu-id="2a6e7-119">Создание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-119">Compose</span></span><br><span data-ttu-id="2a6e7-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-120">Read</span></span> | [<span data-ttu-id="2a6e7-121">Context</span><span class="sxs-lookup"><span data-stu-id="2a6e7-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="2a6e7-122">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2a6e7-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="2a6e7-123">Enumerations</span></span>

| <span data-ttu-id="2a6e7-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="2a6e7-124">Enumeration</span></span> | <span data-ttu-id="2a6e7-125">Способов</span><span class="sxs-lookup"><span data-stu-id="2a6e7-125">Modes</span></span> | <span data-ttu-id="2a6e7-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="2a6e7-126">Return type</span></span> | <span data-ttu-id="2a6e7-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="2a6e7-127">Minimum</span></span><br><span data-ttu-id="2a6e7-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="2a6e7-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2a6e7-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2a6e7-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2a6e7-130">Создание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-130">Compose</span></span><br><span data-ttu-id="2a6e7-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-131">Read</span></span> | <span data-ttu-id="2a6e7-132">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-132">String</span></span> | [<span data-ttu-id="2a6e7-133">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2a6e7-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2a6e7-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2a6e7-135">Создание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-135">Compose</span></span><br><span data-ttu-id="2a6e7-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-136">Read</span></span> | <span data-ttu-id="2a6e7-137">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-137">String</span></span> | [<span data-ttu-id="2a6e7-138">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2a6e7-139">EventType</span><span class="sxs-lookup"><span data-stu-id="2a6e7-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2a6e7-140">Создание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-140">Compose</span></span><br><span data-ttu-id="2a6e7-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-141">Read</span></span> | <span data-ttu-id="2a6e7-142">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-142">String</span></span> | [<span data-ttu-id="2a6e7-143">1,5</span><span class="sxs-lookup"><span data-stu-id="2a6e7-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2a6e7-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2a6e7-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2a6e7-145">Создание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-145">Compose</span></span><br><span data-ttu-id="2a6e7-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-146">Read</span></span> | <span data-ttu-id="2a6e7-147">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-147">String</span></span> | [<span data-ttu-id="2a6e7-148">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2a6e7-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="2a6e7-149">Namespaces</span></span>

<span data-ttu-id="2a6e7-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2a6e7-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="2a6e7-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2a6e7-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="2a6e7-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="2a6e7-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2a6e7-154">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-154">Type</span></span>

*   <span data-ttu-id="2a6e7-155">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a6e7-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a6e7-156">Properties:</span></span>

|<span data-ttu-id="2a6e7-157">Имя</span><span class="sxs-lookup"><span data-stu-id="2a6e7-157">Name</span></span>| <span data-ttu-id="2a6e7-158">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-158">Type</span></span>| <span data-ttu-id="2a6e7-159">Описание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2a6e7-160">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-160">String</span></span>|<span data-ttu-id="2a6e7-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2a6e7-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="2a6e7-162">String</span></span>|<span data-ttu-id="2a6e7-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a6e7-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a6e7-164">Requirements</span></span>

|<span data-ttu-id="2a6e7-165">Требование</span><span class="sxs-lookup"><span data-stu-id="2a6e7-165">Requirement</span></span>| <span data-ttu-id="2a6e7-166">Значение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a6e7-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a6e7-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a6e7-168">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-168">1.1</span></span>|
|[<span data-ttu-id="2a6e7-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a6e7-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a6e7-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2a6e7-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="2a6e7-171">CoercionType: String</span></span>

<span data-ttu-id="2a6e7-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2a6e7-173">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-173">Type</span></span>

*   <span data-ttu-id="2a6e7-174">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a6e7-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a6e7-175">Properties:</span></span>

|<span data-ttu-id="2a6e7-176">Имя</span><span class="sxs-lookup"><span data-stu-id="2a6e7-176">Name</span></span>| <span data-ttu-id="2a6e7-177">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-177">Type</span></span>| <span data-ttu-id="2a6e7-178">Описание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2a6e7-179">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-179">String</span></span>|<span data-ttu-id="2a6e7-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2a6e7-181">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-181">String</span></span>|<span data-ttu-id="2a6e7-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a6e7-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a6e7-183">Requirements</span></span>

|<span data-ttu-id="2a6e7-184">Требование</span><span class="sxs-lookup"><span data-stu-id="2a6e7-184">Requirement</span></span>| <span data-ttu-id="2a6e7-185">Значение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a6e7-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a6e7-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a6e7-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-187">1.1</span></span>|
|[<span data-ttu-id="2a6e7-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a6e7-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a6e7-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2a6e7-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="2a6e7-190">EventType: String</span></span>

<span data-ttu-id="2a6e7-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2a6e7-192">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-192">Type</span></span>

*   <span data-ttu-id="2a6e7-193">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a6e7-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a6e7-194">Properties:</span></span>

| <span data-ttu-id="2a6e7-195">Имя</span><span class="sxs-lookup"><span data-stu-id="2a6e7-195">Name</span></span> | <span data-ttu-id="2a6e7-196">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-196">Type</span></span> | <span data-ttu-id="2a6e7-197">Описание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-197">Description</span></span> | <span data-ttu-id="2a6e7-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="2a6e7-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="2a6e7-199">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-199">String</span></span> | <span data-ttu-id="2a6e7-200">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="2a6e7-201">1.7</span><span class="sxs-lookup"><span data-stu-id="2a6e7-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="2a6e7-202">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-202">String</span></span> | <span data-ttu-id="2a6e7-203">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2a6e7-204">1.5</span><span class="sxs-lookup"><span data-stu-id="2a6e7-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="2a6e7-205">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-205">String</span></span> | <span data-ttu-id="2a6e7-206">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="2a6e7-207">1.7</span><span class="sxs-lookup"><span data-stu-id="2a6e7-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="2a6e7-208">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-208">String</span></span> | <span data-ttu-id="2a6e7-209">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="2a6e7-210">1.7</span><span class="sxs-lookup"><span data-stu-id="2a6e7-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2a6e7-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a6e7-211">Requirements</span></span>

|<span data-ttu-id="2a6e7-212">Требование</span><span class="sxs-lookup"><span data-stu-id="2a6e7-212">Requirement</span></span>| <span data-ttu-id="2a6e7-213">Значение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a6e7-214">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2a6e7-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a6e7-215">1.5</span><span class="sxs-lookup"><span data-stu-id="2a6e7-215">1.5</span></span> |
|[<span data-ttu-id="2a6e7-216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a6e7-216">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a6e7-217">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2a6e7-218">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="2a6e7-218">SourceProperty: String</span></span>

<span data-ttu-id="2a6e7-219">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2a6e7-220">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-220">Type</span></span>

*   <span data-ttu-id="2a6e7-221">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a6e7-222">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a6e7-222">Properties:</span></span>

|<span data-ttu-id="2a6e7-223">Имя</span><span class="sxs-lookup"><span data-stu-id="2a6e7-223">Name</span></span>| <span data-ttu-id="2a6e7-224">Тип</span><span class="sxs-lookup"><span data-stu-id="2a6e7-224">Type</span></span>| <span data-ttu-id="2a6e7-225">Описание</span><span class="sxs-lookup"><span data-stu-id="2a6e7-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2a6e7-226">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-226">String</span></span>|<span data-ttu-id="2a6e7-227">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2a6e7-228">String</span><span class="sxs-lookup"><span data-stu-id="2a6e7-228">String</span></span>|<span data-ttu-id="2a6e7-229">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="2a6e7-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a6e7-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a6e7-230">Requirements</span></span>

|<span data-ttu-id="2a6e7-231">Требование</span><span class="sxs-lookup"><span data-stu-id="2a6e7-231">Requirement</span></span>| <span data-ttu-id="2a6e7-232">Значение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a6e7-233">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a6e7-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a6e7-234">1.1</span><span class="sxs-lookup"><span data-stu-id="2a6e7-234">1.1</span></span>|
|[<span data-ttu-id="2a6e7-235">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a6e7-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a6e7-236">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a6e7-236">Compose or Read</span></span>|
