---
title: Office пространства имен — набор требований 1.9
description: Office пространства имен, доступных для Outlook надстройки с помощью API почтовых ящиков, установленного 1.9.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 203b901c619e19a8e5b9255e36274e2f6e1d1658
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590948"
---
# <a name="office-mailbox-requirement-set-19"></a><span data-ttu-id="bb72f-103">Office (набор требований к почтовым ящикам 1.9)</span><span class="sxs-lookup"><span data-stu-id="bb72f-103">Office (Mailbox requirement set 1.9)</span></span>

<span data-ttu-id="bb72f-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="bb72f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bb72f-106">Требования</span><span class="sxs-lookup"><span data-stu-id="bb72f-106">Requirements</span></span>

|<span data-ttu-id="bb72f-107">Требование</span><span class="sxs-lookup"><span data-stu-id="bb72f-107">Requirement</span></span>| <span data-ttu-id="bb72f-108">Значение</span><span class="sxs-lookup"><span data-stu-id="bb72f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bb72f-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bb72f-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bb72f-110">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-110">1.1</span></span>|
|[<span data-ttu-id="bb72f-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bb72f-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bb72f-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="bb72f-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="bb72f-113">Properties</span></span>

| <span data-ttu-id="bb72f-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="bb72f-114">Property</span></span> | <span data-ttu-id="bb72f-115">Режимы</span><span class="sxs-lookup"><span data-stu-id="bb72f-115">Modes</span></span> | <span data-ttu-id="bb72f-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="bb72f-116">Return type</span></span> | <span data-ttu-id="bb72f-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="bb72f-117">Minimum</span></span><br><span data-ttu-id="bb72f-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="bb72f-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bb72f-119">контекст</span><span class="sxs-lookup"><span data-stu-id="bb72f-119">context</span></span>](office.context.md) | <span data-ttu-id="bb72f-120">Создание</span><span class="sxs-lookup"><span data-stu-id="bb72f-120">Compose</span></span><br><span data-ttu-id="bb72f-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-121">Read</span></span> | [<span data-ttu-id="bb72f-122">Context</span><span class="sxs-lookup"><span data-stu-id="bb72f-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="bb72f-123">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="bb72f-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="bb72f-124">Enumerations</span></span>

| <span data-ttu-id="bb72f-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="bb72f-125">Enumeration</span></span> | <span data-ttu-id="bb72f-126">Режимы</span><span class="sxs-lookup"><span data-stu-id="bb72f-126">Modes</span></span> | <span data-ttu-id="bb72f-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="bb72f-127">Return type</span></span> | <span data-ttu-id="bb72f-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="bb72f-128">Minimum</span></span><br><span data-ttu-id="bb72f-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="bb72f-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bb72f-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bb72f-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bb72f-131">Создание</span><span class="sxs-lookup"><span data-stu-id="bb72f-131">Compose</span></span><br><span data-ttu-id="bb72f-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-132">Read</span></span> | <span data-ttu-id="bb72f-133">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-133">String</span></span> | [<span data-ttu-id="bb72f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bb72f-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bb72f-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bb72f-136">Создание</span><span class="sxs-lookup"><span data-stu-id="bb72f-136">Compose</span></span><br><span data-ttu-id="bb72f-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-137">Read</span></span> | <span data-ttu-id="bb72f-138">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-138">String</span></span> | [<span data-ttu-id="bb72f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bb72f-140">EventType</span><span class="sxs-lookup"><span data-stu-id="bb72f-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="bb72f-141">Создание</span><span class="sxs-lookup"><span data-stu-id="bb72f-141">Compose</span></span><br><span data-ttu-id="bb72f-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-142">Read</span></span> | <span data-ttu-id="bb72f-143">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-143">String</span></span> | [<span data-ttu-id="bb72f-144">1.5</span><span class="sxs-lookup"><span data-stu-id="bb72f-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="bb72f-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bb72f-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bb72f-146">Создание</span><span class="sxs-lookup"><span data-stu-id="bb72f-146">Compose</span></span><br><span data-ttu-id="bb72f-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-147">Read</span></span> | <span data-ttu-id="bb72f-148">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-148">String</span></span> | [<span data-ttu-id="bb72f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="bb72f-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="bb72f-150">Namespaces</span></span>

<span data-ttu-id="bb72f-151">[MailboxEnums:](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true)включает ряд Outlook определенных списков, например , , `ItemType` `EntityType` , `AttachmentType` , , , `RecipientType` и `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="bb72f-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="bb72f-152">Сведения о переумериях</span><span class="sxs-lookup"><span data-stu-id="bb72f-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="bb72f-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="bb72f-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="bb72f-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="bb72f-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bb72f-155">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-155">Type</span></span>

*   <span data-ttu-id="bb72f-156">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bb72f-157">Свойства</span><span class="sxs-lookup"><span data-stu-id="bb72f-157">Properties</span></span>

|<span data-ttu-id="bb72f-158">Имя</span><span class="sxs-lookup"><span data-stu-id="bb72f-158">Name</span></span>| <span data-ttu-id="bb72f-159">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-159">Type</span></span>| <span data-ttu-id="bb72f-160">Описание</span><span class="sxs-lookup"><span data-stu-id="bb72f-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bb72f-161">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-161">String</span></span>|<span data-ttu-id="bb72f-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="bb72f-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bb72f-163">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-163">String</span></span>|<span data-ttu-id="bb72f-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="bb72f-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bb72f-165">Требования</span><span class="sxs-lookup"><span data-stu-id="bb72f-165">Requirements</span></span>

|<span data-ttu-id="bb72f-166">Требование</span><span class="sxs-lookup"><span data-stu-id="bb72f-166">Requirement</span></span>| <span data-ttu-id="bb72f-167">Значение</span><span class="sxs-lookup"><span data-stu-id="bb72f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="bb72f-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bb72f-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bb72f-169">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-169">1.1</span></span>|
|[<span data-ttu-id="bb72f-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bb72f-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bb72f-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="bb72f-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="bb72f-172">CoercionType: String</span></span>

<span data-ttu-id="bb72f-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="bb72f-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bb72f-174">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-174">Type</span></span>

*   <span data-ttu-id="bb72f-175">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bb72f-176">Свойства</span><span class="sxs-lookup"><span data-stu-id="bb72f-176">Properties</span></span>

|<span data-ttu-id="bb72f-177">Имя</span><span class="sxs-lookup"><span data-stu-id="bb72f-177">Name</span></span>| <span data-ttu-id="bb72f-178">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-178">Type</span></span>| <span data-ttu-id="bb72f-179">Описание</span><span class="sxs-lookup"><span data-stu-id="bb72f-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bb72f-180">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-180">String</span></span>|<span data-ttu-id="bb72f-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="bb72f-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bb72f-182">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-182">String</span></span>|<span data-ttu-id="bb72f-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="bb72f-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bb72f-184">Требования</span><span class="sxs-lookup"><span data-stu-id="bb72f-184">Requirements</span></span>

|<span data-ttu-id="bb72f-185">Требование</span><span class="sxs-lookup"><span data-stu-id="bb72f-185">Requirement</span></span>| <span data-ttu-id="bb72f-186">Значение</span><span class="sxs-lookup"><span data-stu-id="bb72f-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="bb72f-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bb72f-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bb72f-188">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-188">1.1</span></span>|
|[<span data-ttu-id="bb72f-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bb72f-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bb72f-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="bb72f-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="bb72f-191">EventType: String</span></span>

<span data-ttu-id="bb72f-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="bb72f-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="bb72f-193">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-193">Type</span></span>

*   <span data-ttu-id="bb72f-194">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bb72f-195">Свойства</span><span class="sxs-lookup"><span data-stu-id="bb72f-195">Properties</span></span>

| <span data-ttu-id="bb72f-196">Имя</span><span class="sxs-lookup"><span data-stu-id="bb72f-196">Name</span></span> | <span data-ttu-id="bb72f-197">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-197">Type</span></span> | <span data-ttu-id="bb72f-198">Описание</span><span class="sxs-lookup"><span data-stu-id="bb72f-198">Description</span></span> | <span data-ttu-id="bb72f-199">Минимальный набор требований</span><span class="sxs-lookup"><span data-stu-id="bb72f-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="bb72f-200">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-200">String</span></span> | <span data-ttu-id="bb72f-201">Изменилась дата или время выбранной встречи или серии.</span><span class="sxs-lookup"><span data-stu-id="bb72f-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="bb72f-202">1.7</span><span class="sxs-lookup"><span data-stu-id="bb72f-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="bb72f-203">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-203">String</span></span> | <span data-ttu-id="bb72f-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="bb72f-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="bb72f-205">1.8</span><span class="sxs-lookup"><span data-stu-id="bb72f-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="bb72f-206">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-206">String</span></span> | <span data-ttu-id="bb72f-207">Расположение выбранного назначения изменилось.</span><span class="sxs-lookup"><span data-stu-id="bb72f-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="bb72f-208">1.8</span><span class="sxs-lookup"><span data-stu-id="bb72f-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="bb72f-209">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-209">String</span></span> | <span data-ttu-id="bb72f-210">Другой элемент Outlook для просмотра при закреплении области задач.</span><span class="sxs-lookup"><span data-stu-id="bb72f-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="bb72f-211">1.5</span><span class="sxs-lookup"><span data-stu-id="bb72f-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="bb72f-212">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-212">String</span></span> | <span data-ttu-id="bb72f-213">Список получателей выбранного элемента или расположения встречи изменен.</span><span class="sxs-lookup"><span data-stu-id="bb72f-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="bb72f-214">1.7</span><span class="sxs-lookup"><span data-stu-id="bb72f-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="bb72f-215">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-215">String</span></span> | <span data-ttu-id="bb72f-216">Изменился шаблон повторяемости выбранной серии.</span><span class="sxs-lookup"><span data-stu-id="bb72f-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="bb72f-217">1.7</span><span class="sxs-lookup"><span data-stu-id="bb72f-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bb72f-218">Требования</span><span class="sxs-lookup"><span data-stu-id="bb72f-218">Requirements</span></span>

|<span data-ttu-id="bb72f-219">Требование</span><span class="sxs-lookup"><span data-stu-id="bb72f-219">Requirement</span></span>| <span data-ttu-id="bb72f-220">Значение</span><span class="sxs-lookup"><span data-stu-id="bb72f-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="bb72f-221">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="bb72f-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bb72f-222">1.5</span><span class="sxs-lookup"><span data-stu-id="bb72f-222">1.5</span></span> |
|[<span data-ttu-id="bb72f-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bb72f-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bb72f-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-224">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="bb72f-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="bb72f-225">SourceProperty: String</span></span>

<span data-ttu-id="bb72f-226">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="bb72f-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bb72f-227">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-227">Type</span></span>

*   <span data-ttu-id="bb72f-228">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bb72f-229">Свойства</span><span class="sxs-lookup"><span data-stu-id="bb72f-229">Properties</span></span>

|<span data-ttu-id="bb72f-230">Имя</span><span class="sxs-lookup"><span data-stu-id="bb72f-230">Name</span></span>| <span data-ttu-id="bb72f-231">Тип</span><span class="sxs-lookup"><span data-stu-id="bb72f-231">Type</span></span>| <span data-ttu-id="bb72f-232">Описание</span><span class="sxs-lookup"><span data-stu-id="bb72f-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bb72f-233">Строка</span><span class="sxs-lookup"><span data-stu-id="bb72f-233">String</span></span>|<span data-ttu-id="bb72f-234">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="bb72f-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bb72f-235">String</span><span class="sxs-lookup"><span data-stu-id="bb72f-235">String</span></span>|<span data-ttu-id="bb72f-236">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="bb72f-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bb72f-237">Требования</span><span class="sxs-lookup"><span data-stu-id="bb72f-237">Requirements</span></span>

|<span data-ttu-id="bb72f-238">Требование</span><span class="sxs-lookup"><span data-stu-id="bb72f-238">Requirement</span></span>| <span data-ttu-id="bb72f-239">Значение</span><span class="sxs-lookup"><span data-stu-id="bb72f-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="bb72f-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bb72f-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bb72f-241">1.1</span><span class="sxs-lookup"><span data-stu-id="bb72f-241">1.1</span></span>|
|[<span data-ttu-id="bb72f-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bb72f-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bb72f-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bb72f-243">Compose or Read</span></span>|
