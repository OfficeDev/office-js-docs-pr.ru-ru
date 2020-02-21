---
title: Пространство имен Office — Предварительная версия набора требований
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2cd04cc6d333439a679803e39357e4d19c550f95
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165512"
---
# <a name="office"></a><span data-ttu-id="348fd-102">Office</span><span class="sxs-lookup"><span data-stu-id="348fd-102">Office</span></span>

<span data-ttu-id="348fd-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="348fd-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="348fd-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="348fd-105">Requirements</span></span>

|<span data-ttu-id="348fd-106">Требование</span><span class="sxs-lookup"><span data-stu-id="348fd-106">Requirement</span></span>| <span data-ttu-id="348fd-107">Значение</span><span class="sxs-lookup"><span data-stu-id="348fd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="348fd-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="348fd-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="348fd-109">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-109">1.1</span></span>|
|[<span data-ttu-id="348fd-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="348fd-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="348fd-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="348fd-112">Properties</span><span class="sxs-lookup"><span data-stu-id="348fd-112">Properties</span></span>

| <span data-ttu-id="348fd-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="348fd-113">Property</span></span> | <span data-ttu-id="348fd-114">Способов</span><span class="sxs-lookup"><span data-stu-id="348fd-114">Modes</span></span> | <span data-ttu-id="348fd-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="348fd-115">Return type</span></span> | <span data-ttu-id="348fd-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="348fd-116">Minimum</span></span><br><span data-ttu-id="348fd-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="348fd-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="348fd-118">контекст</span><span class="sxs-lookup"><span data-stu-id="348fd-118">context</span></span>](office.context.md) | <span data-ttu-id="348fd-119">Создание</span><span class="sxs-lookup"><span data-stu-id="348fd-119">Compose</span></span><br><span data-ttu-id="348fd-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-120">Read</span></span> | [<span data-ttu-id="348fd-121">Context</span><span class="sxs-lookup"><span data-stu-id="348fd-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="348fd-122">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="348fd-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="348fd-123">Enumerations</span></span>

| <span data-ttu-id="348fd-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="348fd-124">Enumeration</span></span> | <span data-ttu-id="348fd-125">Способов</span><span class="sxs-lookup"><span data-stu-id="348fd-125">Modes</span></span> | <span data-ttu-id="348fd-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="348fd-126">Return type</span></span> | <span data-ttu-id="348fd-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="348fd-127">Minimum</span></span><br><span data-ttu-id="348fd-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="348fd-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="348fd-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="348fd-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="348fd-130">Создание</span><span class="sxs-lookup"><span data-stu-id="348fd-130">Compose</span></span><br><span data-ttu-id="348fd-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-131">Read</span></span> | <span data-ttu-id="348fd-132">String</span><span class="sxs-lookup"><span data-stu-id="348fd-132">String</span></span> | [<span data-ttu-id="348fd-133">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="348fd-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="348fd-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="348fd-135">Создание</span><span class="sxs-lookup"><span data-stu-id="348fd-135">Compose</span></span><br><span data-ttu-id="348fd-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-136">Read</span></span> | <span data-ttu-id="348fd-137">String</span><span class="sxs-lookup"><span data-stu-id="348fd-137">String</span></span> | [<span data-ttu-id="348fd-138">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="348fd-139">EventType</span><span class="sxs-lookup"><span data-stu-id="348fd-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="348fd-140">Создание</span><span class="sxs-lookup"><span data-stu-id="348fd-140">Compose</span></span><br><span data-ttu-id="348fd-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-141">Read</span></span> | <span data-ttu-id="348fd-142">String</span><span class="sxs-lookup"><span data-stu-id="348fd-142">String</span></span> | [<span data-ttu-id="348fd-143">1,5</span><span class="sxs-lookup"><span data-stu-id="348fd-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="348fd-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="348fd-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="348fd-145">Создание</span><span class="sxs-lookup"><span data-stu-id="348fd-145">Compose</span></span><br><span data-ttu-id="348fd-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-146">Read</span></span> | <span data-ttu-id="348fd-147">String</span><span class="sxs-lookup"><span data-stu-id="348fd-147">String</span></span> | [<span data-ttu-id="348fd-148">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="348fd-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="348fd-149">Namespaces</span></span>

<span data-ttu-id="348fd-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="348fd-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="348fd-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="348fd-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="348fd-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="348fd-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="348fd-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="348fd-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="348fd-154">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-154">Type</span></span>

*   <span data-ttu-id="348fd-155">String</span><span class="sxs-lookup"><span data-stu-id="348fd-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="348fd-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="348fd-156">Properties:</span></span>

|<span data-ttu-id="348fd-157">Имя</span><span class="sxs-lookup"><span data-stu-id="348fd-157">Name</span></span>| <span data-ttu-id="348fd-158">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-158">Type</span></span>| <span data-ttu-id="348fd-159">Описание</span><span class="sxs-lookup"><span data-stu-id="348fd-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="348fd-160">String</span><span class="sxs-lookup"><span data-stu-id="348fd-160">String</span></span>|<span data-ttu-id="348fd-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="348fd-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="348fd-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="348fd-162">String</span></span>|<span data-ttu-id="348fd-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="348fd-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="348fd-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="348fd-164">Requirements</span></span>

|<span data-ttu-id="348fd-165">Требование</span><span class="sxs-lookup"><span data-stu-id="348fd-165">Requirement</span></span>| <span data-ttu-id="348fd-166">Значение</span><span class="sxs-lookup"><span data-stu-id="348fd-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="348fd-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="348fd-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="348fd-168">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-168">1.1</span></span>|
|[<span data-ttu-id="348fd-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="348fd-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="348fd-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="348fd-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="348fd-171">CoercionType: String</span></span>

<span data-ttu-id="348fd-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="348fd-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="348fd-173">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-173">Type</span></span>

*   <span data-ttu-id="348fd-174">String</span><span class="sxs-lookup"><span data-stu-id="348fd-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="348fd-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="348fd-175">Properties:</span></span>

|<span data-ttu-id="348fd-176">Имя</span><span class="sxs-lookup"><span data-stu-id="348fd-176">Name</span></span>| <span data-ttu-id="348fd-177">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-177">Type</span></span>| <span data-ttu-id="348fd-178">Описание</span><span class="sxs-lookup"><span data-stu-id="348fd-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="348fd-179">String</span><span class="sxs-lookup"><span data-stu-id="348fd-179">String</span></span>|<span data-ttu-id="348fd-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="348fd-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="348fd-181">String</span><span class="sxs-lookup"><span data-stu-id="348fd-181">String</span></span>|<span data-ttu-id="348fd-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="348fd-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="348fd-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="348fd-183">Requirements</span></span>

|<span data-ttu-id="348fd-184">Требование</span><span class="sxs-lookup"><span data-stu-id="348fd-184">Requirement</span></span>| <span data-ttu-id="348fd-185">Значение</span><span class="sxs-lookup"><span data-stu-id="348fd-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="348fd-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="348fd-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="348fd-187">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-187">1.1</span></span>|
|[<span data-ttu-id="348fd-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="348fd-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="348fd-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="348fd-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="348fd-190">EventType: String</span></span>

<span data-ttu-id="348fd-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="348fd-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="348fd-192">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-192">Type</span></span>

*   <span data-ttu-id="348fd-193">String</span><span class="sxs-lookup"><span data-stu-id="348fd-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="348fd-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="348fd-194">Properties:</span></span>

| <span data-ttu-id="348fd-195">Имя</span><span class="sxs-lookup"><span data-stu-id="348fd-195">Name</span></span> | <span data-ttu-id="348fd-196">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-196">Type</span></span> | <span data-ttu-id="348fd-197">Описание</span><span class="sxs-lookup"><span data-stu-id="348fd-197">Description</span></span> | <span data-ttu-id="348fd-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="348fd-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="348fd-199">String</span><span class="sxs-lookup"><span data-stu-id="348fd-199">String</span></span> | <span data-ttu-id="348fd-200">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="348fd-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="348fd-201">1.7</span><span class="sxs-lookup"><span data-stu-id="348fd-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="348fd-202">String</span><span class="sxs-lookup"><span data-stu-id="348fd-202">String</span></span> | <span data-ttu-id="348fd-203">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="348fd-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="348fd-204">1.8</span><span class="sxs-lookup"><span data-stu-id="348fd-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="348fd-205">String</span><span class="sxs-lookup"><span data-stu-id="348fd-205">String</span></span> | <span data-ttu-id="348fd-206">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="348fd-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="348fd-207">1.8</span><span class="sxs-lookup"><span data-stu-id="348fd-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="348fd-208">String</span><span class="sxs-lookup"><span data-stu-id="348fd-208">String</span></span> | <span data-ttu-id="348fd-209">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="348fd-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="348fd-210">1.5</span><span class="sxs-lookup"><span data-stu-id="348fd-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="348fd-211">String</span><span class="sxs-lookup"><span data-stu-id="348fd-211">String</span></span> | <span data-ttu-id="348fd-212">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="348fd-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="348fd-213">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="348fd-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="348fd-214">String</span><span class="sxs-lookup"><span data-stu-id="348fd-214">String</span></span> | <span data-ttu-id="348fd-215">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="348fd-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="348fd-216">1.7</span><span class="sxs-lookup"><span data-stu-id="348fd-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="348fd-217">String</span><span class="sxs-lookup"><span data-stu-id="348fd-217">String</span></span> | <span data-ttu-id="348fd-218">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="348fd-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="348fd-219">1.7</span><span class="sxs-lookup"><span data-stu-id="348fd-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="348fd-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="348fd-220">Requirements</span></span>

|<span data-ttu-id="348fd-221">Требование</span><span class="sxs-lookup"><span data-stu-id="348fd-221">Requirement</span></span>| <span data-ttu-id="348fd-222">Значение</span><span class="sxs-lookup"><span data-stu-id="348fd-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="348fd-223">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="348fd-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="348fd-224">1.5</span><span class="sxs-lookup"><span data-stu-id="348fd-224">1.5</span></span> |
|[<span data-ttu-id="348fd-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="348fd-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="348fd-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-226">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="348fd-227">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="348fd-227">SourceProperty: String</span></span>

<span data-ttu-id="348fd-228">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="348fd-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="348fd-229">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-229">Type</span></span>

*   <span data-ttu-id="348fd-230">String</span><span class="sxs-lookup"><span data-stu-id="348fd-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="348fd-231">Свойства:</span><span class="sxs-lookup"><span data-stu-id="348fd-231">Properties:</span></span>

|<span data-ttu-id="348fd-232">Имя</span><span class="sxs-lookup"><span data-stu-id="348fd-232">Name</span></span>| <span data-ttu-id="348fd-233">Тип</span><span class="sxs-lookup"><span data-stu-id="348fd-233">Type</span></span>| <span data-ttu-id="348fd-234">Описание</span><span class="sxs-lookup"><span data-stu-id="348fd-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="348fd-235">String</span><span class="sxs-lookup"><span data-stu-id="348fd-235">String</span></span>|<span data-ttu-id="348fd-236">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="348fd-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="348fd-237">String</span><span class="sxs-lookup"><span data-stu-id="348fd-237">String</span></span>|<span data-ttu-id="348fd-238">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="348fd-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="348fd-239">Requirements</span><span class="sxs-lookup"><span data-stu-id="348fd-239">Requirements</span></span>

|<span data-ttu-id="348fd-240">Требование</span><span class="sxs-lookup"><span data-stu-id="348fd-240">Requirement</span></span>| <span data-ttu-id="348fd-241">Значение</span><span class="sxs-lookup"><span data-stu-id="348fd-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="348fd-242">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="348fd-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="348fd-243">1.1</span><span class="sxs-lookup"><span data-stu-id="348fd-243">1.1</span></span>|
|[<span data-ttu-id="348fd-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="348fd-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="348fd-245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="348fd-245">Compose or Read</span></span>|
