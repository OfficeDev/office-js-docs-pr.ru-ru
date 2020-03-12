---
title: Пространство имен Office — набор обязательных элементов 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7cc8e6acc60c28b44ec7a2b91bb5e388b2618a31
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554726"
---
# <a name="office"></a><span data-ttu-id="f9da7-102">Office</span><span class="sxs-lookup"><span data-stu-id="f9da7-102">Office</span></span>

<span data-ttu-id="f9da7-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f9da7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9da7-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="f9da7-105">Requirements</span></span>

|<span data-ttu-id="f9da7-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f9da7-106">Requirement</span></span>| <span data-ttu-id="f9da7-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f9da7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9da7-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f9da7-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f9da7-109">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-109">1.1</span></span>|
|[<span data-ttu-id="f9da7-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f9da7-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f9da7-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f9da7-112">Properties</span><span class="sxs-lookup"><span data-stu-id="f9da7-112">Properties</span></span>

| <span data-ttu-id="f9da7-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="f9da7-113">Property</span></span> | <span data-ttu-id="f9da7-114">Способов</span><span class="sxs-lookup"><span data-stu-id="f9da7-114">Modes</span></span> | <span data-ttu-id="f9da7-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="f9da7-115">Return type</span></span> | <span data-ttu-id="f9da7-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f9da7-116">Minimum</span></span><br><span data-ttu-id="f9da7-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="f9da7-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f9da7-118">контекст</span><span class="sxs-lookup"><span data-stu-id="f9da7-118">context</span></span>](office.context.md) | <span data-ttu-id="f9da7-119">Создание</span><span class="sxs-lookup"><span data-stu-id="f9da7-119">Compose</span></span><br><span data-ttu-id="f9da7-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-120">Read</span></span> | [<span data-ttu-id="f9da7-121">Context</span><span class="sxs-lookup"><span data-stu-id="f9da7-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="f9da7-122">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f9da7-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="f9da7-123">Enumerations</span></span>

| <span data-ttu-id="f9da7-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="f9da7-124">Enumeration</span></span> | <span data-ttu-id="f9da7-125">Способов</span><span class="sxs-lookup"><span data-stu-id="f9da7-125">Modes</span></span> | <span data-ttu-id="f9da7-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="f9da7-126">Return type</span></span> | <span data-ttu-id="f9da7-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f9da7-127">Minimum</span></span><br><span data-ttu-id="f9da7-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="f9da7-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f9da7-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f9da7-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f9da7-130">Создание</span><span class="sxs-lookup"><span data-stu-id="f9da7-130">Compose</span></span><br><span data-ttu-id="f9da7-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-131">Read</span></span> | <span data-ttu-id="f9da7-132">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-132">String</span></span> | [<span data-ttu-id="f9da7-133">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f9da7-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f9da7-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f9da7-135">Создание</span><span class="sxs-lookup"><span data-stu-id="f9da7-135">Compose</span></span><br><span data-ttu-id="f9da7-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-136">Read</span></span> | <span data-ttu-id="f9da7-137">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-137">String</span></span> | [<span data-ttu-id="f9da7-138">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f9da7-139">EventType</span><span class="sxs-lookup"><span data-stu-id="f9da7-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f9da7-140">Создание</span><span class="sxs-lookup"><span data-stu-id="f9da7-140">Compose</span></span><br><span data-ttu-id="f9da7-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-141">Read</span></span> | <span data-ttu-id="f9da7-142">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-142">String</span></span> | [<span data-ttu-id="f9da7-143">1,5</span><span class="sxs-lookup"><span data-stu-id="f9da7-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f9da7-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f9da7-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f9da7-145">Создание</span><span class="sxs-lookup"><span data-stu-id="f9da7-145">Compose</span></span><br><span data-ttu-id="f9da7-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-146">Read</span></span> | <span data-ttu-id="f9da7-147">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-147">String</span></span> | [<span data-ttu-id="f9da7-148">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f9da7-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f9da7-149">Namespaces</span></span>

<span data-ttu-id="f9da7-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f9da7-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f9da7-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="f9da7-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f9da7-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="f9da7-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="f9da7-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="f9da7-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f9da7-154">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-154">Type</span></span>

*   <span data-ttu-id="f9da7-155">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9da7-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f9da7-156">Properties:</span></span>

|<span data-ttu-id="f9da7-157">Имя</span><span class="sxs-lookup"><span data-stu-id="f9da7-157">Name</span></span>| <span data-ttu-id="f9da7-158">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-158">Type</span></span>| <span data-ttu-id="f9da7-159">Описание</span><span class="sxs-lookup"><span data-stu-id="f9da7-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f9da7-160">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-160">String</span></span>|<span data-ttu-id="f9da7-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="f9da7-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f9da7-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="f9da7-162">String</span></span>|<span data-ttu-id="f9da7-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="f9da7-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9da7-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="f9da7-164">Requirements</span></span>

|<span data-ttu-id="f9da7-165">Требование</span><span class="sxs-lookup"><span data-stu-id="f9da7-165">Requirement</span></span>| <span data-ttu-id="f9da7-166">Значение</span><span class="sxs-lookup"><span data-stu-id="f9da7-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9da7-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f9da7-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f9da7-168">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-168">1.1</span></span>|
|[<span data-ttu-id="f9da7-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f9da7-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f9da7-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f9da7-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="f9da7-171">CoercionType: String</span></span>

<span data-ttu-id="f9da7-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f9da7-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9da7-173">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-173">Type</span></span>

*   <span data-ttu-id="f9da7-174">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9da7-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f9da7-175">Properties:</span></span>

|<span data-ttu-id="f9da7-176">Имя</span><span class="sxs-lookup"><span data-stu-id="f9da7-176">Name</span></span>| <span data-ttu-id="f9da7-177">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-177">Type</span></span>| <span data-ttu-id="f9da7-178">Описание</span><span class="sxs-lookup"><span data-stu-id="f9da7-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f9da7-179">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-179">String</span></span>|<span data-ttu-id="f9da7-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f9da7-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f9da7-181">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-181">String</span></span>|<span data-ttu-id="f9da7-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="f9da7-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9da7-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="f9da7-183">Requirements</span></span>

|<span data-ttu-id="f9da7-184">Требование</span><span class="sxs-lookup"><span data-stu-id="f9da7-184">Requirement</span></span>| <span data-ttu-id="f9da7-185">Значение</span><span class="sxs-lookup"><span data-stu-id="f9da7-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9da7-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f9da7-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f9da7-187">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-187">1.1</span></span>|
|[<span data-ttu-id="f9da7-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f9da7-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f9da7-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f9da7-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="f9da7-190">EventType: String</span></span>

<span data-ttu-id="f9da7-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="f9da7-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f9da7-192">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-192">Type</span></span>

*   <span data-ttu-id="f9da7-193">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9da7-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f9da7-194">Properties:</span></span>

| <span data-ttu-id="f9da7-195">Имя</span><span class="sxs-lookup"><span data-stu-id="f9da7-195">Name</span></span> | <span data-ttu-id="f9da7-196">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-196">Type</span></span> | <span data-ttu-id="f9da7-197">Описание</span><span class="sxs-lookup"><span data-stu-id="f9da7-197">Description</span></span> | <span data-ttu-id="f9da7-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="f9da7-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="f9da7-199">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-199">String</span></span> | <span data-ttu-id="f9da7-200">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="f9da7-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f9da7-201">1.5</span><span class="sxs-lookup"><span data-stu-id="f9da7-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9da7-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="f9da7-202">Requirements</span></span>

|<span data-ttu-id="f9da7-203">Требование</span><span class="sxs-lookup"><span data-stu-id="f9da7-203">Requirement</span></span>| <span data-ttu-id="f9da7-204">Значение</span><span class="sxs-lookup"><span data-stu-id="f9da7-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9da7-205">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f9da7-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f9da7-206">1.5</span><span class="sxs-lookup"><span data-stu-id="f9da7-206">1.5</span></span> |
|[<span data-ttu-id="f9da7-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f9da7-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f9da7-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f9da7-209">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="f9da7-209">SourceProperty: String</span></span>

<span data-ttu-id="f9da7-210">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f9da7-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9da7-211">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-211">Type</span></span>

*   <span data-ttu-id="f9da7-212">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9da7-213">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f9da7-213">Properties:</span></span>

|<span data-ttu-id="f9da7-214">Имя</span><span class="sxs-lookup"><span data-stu-id="f9da7-214">Name</span></span>| <span data-ttu-id="f9da7-215">Тип</span><span class="sxs-lookup"><span data-stu-id="f9da7-215">Type</span></span>| <span data-ttu-id="f9da7-216">Описание</span><span class="sxs-lookup"><span data-stu-id="f9da7-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f9da7-217">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-217">String</span></span>|<span data-ttu-id="f9da7-218">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="f9da7-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f9da7-219">String</span><span class="sxs-lookup"><span data-stu-id="f9da7-219">String</span></span>|<span data-ttu-id="f9da7-220">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="f9da7-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9da7-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="f9da7-221">Requirements</span></span>

|<span data-ttu-id="f9da7-222">Требование</span><span class="sxs-lookup"><span data-stu-id="f9da7-222">Requirement</span></span>| <span data-ttu-id="f9da7-223">Значение</span><span class="sxs-lookup"><span data-stu-id="f9da7-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9da7-224">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f9da7-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f9da7-225">1.1</span><span class="sxs-lookup"><span data-stu-id="f9da7-225">1.1</span></span>|
|[<span data-ttu-id="f9da7-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f9da7-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f9da7-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f9da7-227">Compose or Read</span></span>|
