---
title: Пространство имен Office — набор обязательных элементов 1,1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: d5fafad5866caa3d8bf8486fec6c8797fc7a7a8f
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815013"
---
# <a name="office"></a><span data-ttu-id="8e46e-102">Office</span><span class="sxs-lookup"><span data-stu-id="8e46e-102">Office</span></span>

<span data-ttu-id="8e46e-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="8e46e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e46e-105">Требования</span><span class="sxs-lookup"><span data-stu-id="8e46e-105">Requirements</span></span>

|<span data-ttu-id="8e46e-106">Требование</span><span class="sxs-lookup"><span data-stu-id="8e46e-106">Requirement</span></span>| <span data-ttu-id="8e46e-107">Значение</span><span class="sxs-lookup"><span data-stu-id="8e46e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e46e-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e46e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e46e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-109">1.1</span></span>|
|[<span data-ttu-id="8e46e-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e46e-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e46e-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8e46e-112">Properties</span><span class="sxs-lookup"><span data-stu-id="8e46e-112">Properties</span></span>

| <span data-ttu-id="8e46e-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="8e46e-113">Property</span></span> | <span data-ttu-id="8e46e-114">Способов</span><span class="sxs-lookup"><span data-stu-id="8e46e-114">Modes</span></span> | <span data-ttu-id="8e46e-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="8e46e-115">Return type</span></span> | <span data-ttu-id="8e46e-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="8e46e-116">Minimum</span></span><br><span data-ttu-id="8e46e-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="8e46e-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8e46e-118">контекст</span><span class="sxs-lookup"><span data-stu-id="8e46e-118">context</span></span>](office.context.md) | <span data-ttu-id="8e46e-119">Создание</span><span class="sxs-lookup"><span data-stu-id="8e46e-119">Compose</span></span><br><span data-ttu-id="8e46e-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-120">Read</span></span> | [<span data-ttu-id="8e46e-121">Context</span><span class="sxs-lookup"><span data-stu-id="8e46e-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="8e46e-122">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8e46e-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="8e46e-123">Enumerations</span></span>

| <span data-ttu-id="8e46e-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="8e46e-124">Enumeration</span></span> | <span data-ttu-id="8e46e-125">Способов</span><span class="sxs-lookup"><span data-stu-id="8e46e-125">Modes</span></span> | <span data-ttu-id="8e46e-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="8e46e-126">Return type</span></span> | <span data-ttu-id="8e46e-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="8e46e-127">Minimum</span></span><br><span data-ttu-id="8e46e-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="8e46e-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8e46e-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8e46e-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8e46e-130">Создание</span><span class="sxs-lookup"><span data-stu-id="8e46e-130">Compose</span></span><br><span data-ttu-id="8e46e-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-131">Read</span></span> | <span data-ttu-id="8e46e-132">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-132">String</span></span> | [<span data-ttu-id="8e46e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e46e-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e46e-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8e46e-135">Создание</span><span class="sxs-lookup"><span data-stu-id="8e46e-135">Compose</span></span><br><span data-ttu-id="8e46e-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-136">Read</span></span> | <span data-ttu-id="8e46e-137">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-137">String</span></span> | [<span data-ttu-id="8e46e-138">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e46e-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8e46e-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8e46e-140">Создание</span><span class="sxs-lookup"><span data-stu-id="8e46e-140">Compose</span></span><br><span data-ttu-id="8e46e-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-141">Read</span></span> | <span data-ttu-id="8e46e-142">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-142">String</span></span> | [<span data-ttu-id="8e46e-143">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8e46e-144">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="8e46e-144">Namespaces</span></span>

<span data-ttu-id="8e46e-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="8e46e-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8e46e-146">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="8e46e-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8e46e-147">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="8e46e-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="8e46e-148">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="8e46e-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8e46e-149">Тип</span><span class="sxs-lookup"><span data-stu-id="8e46e-149">Type</span></span>

*   <span data-ttu-id="8e46e-150">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e46e-151">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e46e-151">Properties:</span></span>

|<span data-ttu-id="8e46e-152">Имя</span><span class="sxs-lookup"><span data-stu-id="8e46e-152">Name</span></span>| <span data-ttu-id="8e46e-153">Тип</span><span class="sxs-lookup"><span data-stu-id="8e46e-153">Type</span></span>| <span data-ttu-id="8e46e-154">Описание</span><span class="sxs-lookup"><span data-stu-id="8e46e-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8e46e-155">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-155">String</span></span>|<span data-ttu-id="8e46e-156">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="8e46e-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8e46e-157">Для указания</span><span class="sxs-lookup"><span data-stu-id="8e46e-157">String</span></span>|<span data-ttu-id="8e46e-158">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="8e46e-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e46e-159">Требования</span><span class="sxs-lookup"><span data-stu-id="8e46e-159">Requirements</span></span>

|<span data-ttu-id="8e46e-160">Требование</span><span class="sxs-lookup"><span data-stu-id="8e46e-160">Requirement</span></span>| <span data-ttu-id="8e46e-161">Значение</span><span class="sxs-lookup"><span data-stu-id="8e46e-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e46e-162">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e46e-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e46e-163">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-163">1.1</span></span>|
|[<span data-ttu-id="8e46e-164">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e46e-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e46e-165">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8e46e-166">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="8e46e-166">CoercionType: String</span></span>

<span data-ttu-id="8e46e-167">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="8e46e-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e46e-168">Тип</span><span class="sxs-lookup"><span data-stu-id="8e46e-168">Type</span></span>

*   <span data-ttu-id="8e46e-169">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e46e-170">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e46e-170">Properties:</span></span>

|<span data-ttu-id="8e46e-171">Имя</span><span class="sxs-lookup"><span data-stu-id="8e46e-171">Name</span></span>| <span data-ttu-id="8e46e-172">Тип</span><span class="sxs-lookup"><span data-stu-id="8e46e-172">Type</span></span>| <span data-ttu-id="8e46e-173">Описание</span><span class="sxs-lookup"><span data-stu-id="8e46e-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8e46e-174">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-174">String</span></span>|<span data-ttu-id="8e46e-175">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="8e46e-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8e46e-176">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-176">String</span></span>|<span data-ttu-id="8e46e-177">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="8e46e-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e46e-178">Требования</span><span class="sxs-lookup"><span data-stu-id="8e46e-178">Requirements</span></span>

|<span data-ttu-id="8e46e-179">Требование</span><span class="sxs-lookup"><span data-stu-id="8e46e-179">Requirement</span></span>| <span data-ttu-id="8e46e-180">Значение</span><span class="sxs-lookup"><span data-stu-id="8e46e-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e46e-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e46e-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e46e-182">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-182">1.1</span></span>|
|[<span data-ttu-id="8e46e-183">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e46e-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e46e-184">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8e46e-185">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="8e46e-185">SourceProperty: String</span></span>

<span data-ttu-id="8e46e-186">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="8e46e-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e46e-187">Тип</span><span class="sxs-lookup"><span data-stu-id="8e46e-187">Type</span></span>

*   <span data-ttu-id="8e46e-188">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e46e-189">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e46e-189">Properties:</span></span>

|<span data-ttu-id="8e46e-190">Имя</span><span class="sxs-lookup"><span data-stu-id="8e46e-190">Name</span></span>| <span data-ttu-id="8e46e-191">Тип</span><span class="sxs-lookup"><span data-stu-id="8e46e-191">Type</span></span>| <span data-ttu-id="8e46e-192">Описание</span><span class="sxs-lookup"><span data-stu-id="8e46e-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8e46e-193">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-193">String</span></span>|<span data-ttu-id="8e46e-194">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="8e46e-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8e46e-195">String</span><span class="sxs-lookup"><span data-stu-id="8e46e-195">String</span></span>|<span data-ttu-id="8e46e-196">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="8e46e-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e46e-197">Требования</span><span class="sxs-lookup"><span data-stu-id="8e46e-197">Requirements</span></span>

|<span data-ttu-id="8e46e-198">Требование</span><span class="sxs-lookup"><span data-stu-id="8e46e-198">Requirement</span></span>| <span data-ttu-id="8e46e-199">Значение</span><span class="sxs-lookup"><span data-stu-id="8e46e-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e46e-200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e46e-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e46e-201">1.1</span><span class="sxs-lookup"><span data-stu-id="8e46e-201">1.1</span></span>|
|[<span data-ttu-id="8e46e-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e46e-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e46e-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e46e-203">Compose or Read</span></span>|
