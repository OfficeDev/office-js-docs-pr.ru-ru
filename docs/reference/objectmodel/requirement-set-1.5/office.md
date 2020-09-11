---
title: Пространство имен Office — набор обязательных элементов 1,5
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 141fd124ba5778a5ae576c7b4cd2c749a9c4bd6f
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430599"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="3ded1-103">Office (набор требований для почтового ящика 1,5)</span><span class="sxs-lookup"><span data-stu-id="3ded1-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="3ded1-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3ded1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ded1-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ded1-106">Requirements</span></span>

|<span data-ttu-id="3ded1-107">Требование</span><span class="sxs-lookup"><span data-stu-id="3ded1-107">Requirement</span></span>| <span data-ttu-id="3ded1-108">Значение</span><span class="sxs-lookup"><span data-stu-id="3ded1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ded1-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3ded1-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ded1-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-110">1.1</span></span>|
|[<span data-ttu-id="3ded1-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3ded1-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ded1-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3ded1-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="3ded1-113">Properties</span></span>

| <span data-ttu-id="3ded1-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="3ded1-114">Property</span></span> | <span data-ttu-id="3ded1-115">Способов</span><span class="sxs-lookup"><span data-stu-id="3ded1-115">Modes</span></span> | <span data-ttu-id="3ded1-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="3ded1-116">Return type</span></span> | <span data-ttu-id="3ded1-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="3ded1-117">Minimum</span></span><br><span data-ttu-id="3ded1-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="3ded1-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3ded1-119">контекст</span><span class="sxs-lookup"><span data-stu-id="3ded1-119">context</span></span>](office.context.md) | <span data-ttu-id="3ded1-120">Создание</span><span class="sxs-lookup"><span data-stu-id="3ded1-120">Compose</span></span><br><span data-ttu-id="3ded1-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-121">Read</span></span> | [<span data-ttu-id="3ded1-122">Context</span><span class="sxs-lookup"><span data-stu-id="3ded1-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="3ded1-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3ded1-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="3ded1-124">Enumerations</span></span>

| <span data-ttu-id="3ded1-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="3ded1-125">Enumeration</span></span> | <span data-ttu-id="3ded1-126">Способов</span><span class="sxs-lookup"><span data-stu-id="3ded1-126">Modes</span></span> | <span data-ttu-id="3ded1-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="3ded1-127">Return type</span></span> | <span data-ttu-id="3ded1-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="3ded1-128">Minimum</span></span><br><span data-ttu-id="3ded1-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="3ded1-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3ded1-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3ded1-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3ded1-131">Создание</span><span class="sxs-lookup"><span data-stu-id="3ded1-131">Compose</span></span><br><span data-ttu-id="3ded1-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-132">Read</span></span> | <span data-ttu-id="3ded1-133">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-133">String</span></span> | [<span data-ttu-id="3ded1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3ded1-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3ded1-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3ded1-136">Создание</span><span class="sxs-lookup"><span data-stu-id="3ded1-136">Compose</span></span><br><span data-ttu-id="3ded1-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-137">Read</span></span> | <span data-ttu-id="3ded1-138">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-138">String</span></span> | [<span data-ttu-id="3ded1-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3ded1-140">EventType</span><span class="sxs-lookup"><span data-stu-id="3ded1-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3ded1-141">Создание</span><span class="sxs-lookup"><span data-stu-id="3ded1-141">Compose</span></span><br><span data-ttu-id="3ded1-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-142">Read</span></span> | <span data-ttu-id="3ded1-143">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-143">String</span></span> | [<span data-ttu-id="3ded1-144">1,5</span><span class="sxs-lookup"><span data-stu-id="3ded1-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3ded1-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3ded1-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3ded1-146">Создание</span><span class="sxs-lookup"><span data-stu-id="3ded1-146">Compose</span></span><br><span data-ttu-id="3ded1-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-147">Read</span></span> | <span data-ttu-id="3ded1-148">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-148">String</span></span> | [<span data-ttu-id="3ded1-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3ded1-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="3ded1-150">Namespaces</span></span>

<span data-ttu-id="3ded1-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="3ded1-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3ded1-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="3ded1-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3ded1-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="3ded1-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="3ded1-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="3ded1-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3ded1-155">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-155">Type</span></span>

*   <span data-ttu-id="3ded1-156">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ded1-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3ded1-157">Properties:</span></span>

|<span data-ttu-id="3ded1-158">Имя</span><span class="sxs-lookup"><span data-stu-id="3ded1-158">Name</span></span>| <span data-ttu-id="3ded1-159">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-159">Type</span></span>| <span data-ttu-id="3ded1-160">Описание</span><span class="sxs-lookup"><span data-stu-id="3ded1-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3ded1-161">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-161">String</span></span>|<span data-ttu-id="3ded1-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="3ded1-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3ded1-163">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-163">String</span></span>|<span data-ttu-id="3ded1-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="3ded1-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ded1-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ded1-165">Requirements</span></span>

|<span data-ttu-id="3ded1-166">Требование</span><span class="sxs-lookup"><span data-stu-id="3ded1-166">Requirement</span></span>| <span data-ttu-id="3ded1-167">Значение</span><span class="sxs-lookup"><span data-stu-id="3ded1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ded1-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3ded1-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ded1-169">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-169">1.1</span></span>|
|[<span data-ttu-id="3ded1-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3ded1-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ded1-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3ded1-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="3ded1-172">CoercionType: String</span></span>

<span data-ttu-id="3ded1-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="3ded1-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3ded1-174">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-174">Type</span></span>

*   <span data-ttu-id="3ded1-175">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ded1-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3ded1-176">Properties:</span></span>

|<span data-ttu-id="3ded1-177">Имя</span><span class="sxs-lookup"><span data-stu-id="3ded1-177">Name</span></span>| <span data-ttu-id="3ded1-178">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-178">Type</span></span>| <span data-ttu-id="3ded1-179">Описание</span><span class="sxs-lookup"><span data-stu-id="3ded1-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3ded1-180">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-180">String</span></span>|<span data-ttu-id="3ded1-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="3ded1-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3ded1-182">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-182">String</span></span>|<span data-ttu-id="3ded1-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="3ded1-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ded1-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ded1-184">Requirements</span></span>

|<span data-ttu-id="3ded1-185">Требование</span><span class="sxs-lookup"><span data-stu-id="3ded1-185">Requirement</span></span>| <span data-ttu-id="3ded1-186">Значение</span><span class="sxs-lookup"><span data-stu-id="3ded1-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ded1-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3ded1-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ded1-188">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-188">1.1</span></span>|
|[<span data-ttu-id="3ded1-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3ded1-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ded1-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3ded1-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="3ded1-191">EventType: String</span></span>

<span data-ttu-id="3ded1-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="3ded1-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3ded1-193">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-193">Type</span></span>

*   <span data-ttu-id="3ded1-194">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ded1-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3ded1-195">Properties:</span></span>

| <span data-ttu-id="3ded1-196">Имя</span><span class="sxs-lookup"><span data-stu-id="3ded1-196">Name</span></span> | <span data-ttu-id="3ded1-197">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-197">Type</span></span> | <span data-ttu-id="3ded1-198">Описание</span><span class="sxs-lookup"><span data-stu-id="3ded1-198">Description</span></span> | <span data-ttu-id="3ded1-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="3ded1-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="3ded1-200">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-200">String</span></span> | <span data-ttu-id="3ded1-201">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="3ded1-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3ded1-202">1.5</span><span class="sxs-lookup"><span data-stu-id="3ded1-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ded1-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ded1-203">Requirements</span></span>

|<span data-ttu-id="3ded1-204">Требование</span><span class="sxs-lookup"><span data-stu-id="3ded1-204">Requirement</span></span>| <span data-ttu-id="3ded1-205">Значение</span><span class="sxs-lookup"><span data-stu-id="3ded1-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ded1-206">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3ded1-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ded1-207">1.5</span><span class="sxs-lookup"><span data-stu-id="3ded1-207">1.5</span></span> |
|[<span data-ttu-id="3ded1-208">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3ded1-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ded1-209">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3ded1-210">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="3ded1-210">SourceProperty: String</span></span>

<span data-ttu-id="3ded1-211">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="3ded1-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3ded1-212">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-212">Type</span></span>

*   <span data-ttu-id="3ded1-213">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ded1-214">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3ded1-214">Properties:</span></span>

|<span data-ttu-id="3ded1-215">Имя</span><span class="sxs-lookup"><span data-stu-id="3ded1-215">Name</span></span>| <span data-ttu-id="3ded1-216">Тип</span><span class="sxs-lookup"><span data-stu-id="3ded1-216">Type</span></span>| <span data-ttu-id="3ded1-217">Описание</span><span class="sxs-lookup"><span data-stu-id="3ded1-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3ded1-218">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-218">String</span></span>|<span data-ttu-id="3ded1-219">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="3ded1-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3ded1-220">String</span><span class="sxs-lookup"><span data-stu-id="3ded1-220">String</span></span>|<span data-ttu-id="3ded1-221">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="3ded1-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ded1-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ded1-222">Requirements</span></span>

|<span data-ttu-id="3ded1-223">Требование</span><span class="sxs-lookup"><span data-stu-id="3ded1-223">Requirement</span></span>| <span data-ttu-id="3ded1-224">Значение</span><span class="sxs-lookup"><span data-stu-id="3ded1-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ded1-225">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3ded1-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ded1-226">1.1</span><span class="sxs-lookup"><span data-stu-id="3ded1-226">1.1</span></span>|
|[<span data-ttu-id="3ded1-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3ded1-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ded1-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3ded1-228">Compose or Read</span></span>|
