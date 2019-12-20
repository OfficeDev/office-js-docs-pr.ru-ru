---
title: Пространство имен Office — набор обязательных элементов 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 714bbd6dfdcd47687a2309c24fd666168aed6556
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814329"
---
# <a name="office"></a><span data-ttu-id="35395-102">Office</span><span class="sxs-lookup"><span data-stu-id="35395-102">Office</span></span>

<span data-ttu-id="35395-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="35395-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="35395-105">Требования</span><span class="sxs-lookup"><span data-stu-id="35395-105">Requirements</span></span>

|<span data-ttu-id="35395-106">Требование</span><span class="sxs-lookup"><span data-stu-id="35395-106">Requirement</span></span>| <span data-ttu-id="35395-107">Значение</span><span class="sxs-lookup"><span data-stu-id="35395-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="35395-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35395-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35395-109">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-109">1.1</span></span>|
|[<span data-ttu-id="35395-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35395-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35395-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35395-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="35395-112">Properties</span><span class="sxs-lookup"><span data-stu-id="35395-112">Properties</span></span>

| <span data-ttu-id="35395-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="35395-113">Property</span></span> | <span data-ttu-id="35395-114">Способов</span><span class="sxs-lookup"><span data-stu-id="35395-114">Modes</span></span> | <span data-ttu-id="35395-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="35395-115">Return type</span></span> | <span data-ttu-id="35395-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="35395-116">Minimum</span></span><br><span data-ttu-id="35395-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="35395-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="35395-118">контекст</span><span class="sxs-lookup"><span data-stu-id="35395-118">context</span></span>](office.context.md) | <span data-ttu-id="35395-119">Создание</span><span class="sxs-lookup"><span data-stu-id="35395-119">Compose</span></span><br><span data-ttu-id="35395-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="35395-120">Read</span></span> | [<span data-ttu-id="35395-121">Context</span><span class="sxs-lookup"><span data-stu-id="35395-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="35395-122">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="35395-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="35395-123">Enumerations</span></span>

| <span data-ttu-id="35395-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="35395-124">Enumeration</span></span> | <span data-ttu-id="35395-125">Способов</span><span class="sxs-lookup"><span data-stu-id="35395-125">Modes</span></span> | <span data-ttu-id="35395-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="35395-126">Return type</span></span> | <span data-ttu-id="35395-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="35395-127">Minimum</span></span><br><span data-ttu-id="35395-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="35395-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="35395-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="35395-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="35395-130">Создание</span><span class="sxs-lookup"><span data-stu-id="35395-130">Compose</span></span><br><span data-ttu-id="35395-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="35395-131">Read</span></span> | <span data-ttu-id="35395-132">String</span><span class="sxs-lookup"><span data-stu-id="35395-132">String</span></span> | [<span data-ttu-id="35395-133">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="35395-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="35395-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="35395-135">Создание</span><span class="sxs-lookup"><span data-stu-id="35395-135">Compose</span></span><br><span data-ttu-id="35395-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="35395-136">Read</span></span> | <span data-ttu-id="35395-137">String</span><span class="sxs-lookup"><span data-stu-id="35395-137">String</span></span> | [<span data-ttu-id="35395-138">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="35395-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="35395-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="35395-140">Создание</span><span class="sxs-lookup"><span data-stu-id="35395-140">Compose</span></span><br><span data-ttu-id="35395-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="35395-141">Read</span></span> | <span data-ttu-id="35395-142">String</span><span class="sxs-lookup"><span data-stu-id="35395-142">String</span></span> | [<span data-ttu-id="35395-143">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="35395-144">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="35395-144">Namespaces</span></span>

<span data-ttu-id="35395-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="35395-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="35395-146">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="35395-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="35395-147">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="35395-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="35395-148">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="35395-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="35395-149">Тип</span><span class="sxs-lookup"><span data-stu-id="35395-149">Type</span></span>

*   <span data-ttu-id="35395-150">String</span><span class="sxs-lookup"><span data-stu-id="35395-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35395-151">Свойства:</span><span class="sxs-lookup"><span data-stu-id="35395-151">Properties:</span></span>

|<span data-ttu-id="35395-152">Имя</span><span class="sxs-lookup"><span data-stu-id="35395-152">Name</span></span>| <span data-ttu-id="35395-153">Тип</span><span class="sxs-lookup"><span data-stu-id="35395-153">Type</span></span>| <span data-ttu-id="35395-154">Описание</span><span class="sxs-lookup"><span data-stu-id="35395-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="35395-155">String</span><span class="sxs-lookup"><span data-stu-id="35395-155">String</span></span>|<span data-ttu-id="35395-156">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="35395-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="35395-157">Для указания</span><span class="sxs-lookup"><span data-stu-id="35395-157">String</span></span>|<span data-ttu-id="35395-158">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="35395-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35395-159">Требования</span><span class="sxs-lookup"><span data-stu-id="35395-159">Requirements</span></span>

|<span data-ttu-id="35395-160">Требование</span><span class="sxs-lookup"><span data-stu-id="35395-160">Requirement</span></span>| <span data-ttu-id="35395-161">Значение</span><span class="sxs-lookup"><span data-stu-id="35395-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="35395-162">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35395-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35395-163">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-163">1.1</span></span>|
|[<span data-ttu-id="35395-164">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35395-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35395-165">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35395-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="35395-166">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="35395-166">CoercionType: String</span></span>

<span data-ttu-id="35395-167">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="35395-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="35395-168">Тип</span><span class="sxs-lookup"><span data-stu-id="35395-168">Type</span></span>

*   <span data-ttu-id="35395-169">String</span><span class="sxs-lookup"><span data-stu-id="35395-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35395-170">Свойства:</span><span class="sxs-lookup"><span data-stu-id="35395-170">Properties:</span></span>

|<span data-ttu-id="35395-171">Имя</span><span class="sxs-lookup"><span data-stu-id="35395-171">Name</span></span>| <span data-ttu-id="35395-172">Тип</span><span class="sxs-lookup"><span data-stu-id="35395-172">Type</span></span>| <span data-ttu-id="35395-173">Описание</span><span class="sxs-lookup"><span data-stu-id="35395-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="35395-174">String</span><span class="sxs-lookup"><span data-stu-id="35395-174">String</span></span>|<span data-ttu-id="35395-175">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="35395-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="35395-176">String</span><span class="sxs-lookup"><span data-stu-id="35395-176">String</span></span>|<span data-ttu-id="35395-177">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="35395-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35395-178">Требования</span><span class="sxs-lookup"><span data-stu-id="35395-178">Requirements</span></span>

|<span data-ttu-id="35395-179">Требование</span><span class="sxs-lookup"><span data-stu-id="35395-179">Requirement</span></span>| <span data-ttu-id="35395-180">Значение</span><span class="sxs-lookup"><span data-stu-id="35395-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="35395-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35395-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35395-182">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-182">1.1</span></span>|
|[<span data-ttu-id="35395-183">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35395-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35395-184">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35395-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="35395-185">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="35395-185">SourceProperty: String</span></span>

<span data-ttu-id="35395-186">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="35395-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="35395-187">Тип</span><span class="sxs-lookup"><span data-stu-id="35395-187">Type</span></span>

*   <span data-ttu-id="35395-188">String</span><span class="sxs-lookup"><span data-stu-id="35395-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35395-189">Свойства:</span><span class="sxs-lookup"><span data-stu-id="35395-189">Properties:</span></span>

|<span data-ttu-id="35395-190">Имя</span><span class="sxs-lookup"><span data-stu-id="35395-190">Name</span></span>| <span data-ttu-id="35395-191">Тип</span><span class="sxs-lookup"><span data-stu-id="35395-191">Type</span></span>| <span data-ttu-id="35395-192">Описание</span><span class="sxs-lookup"><span data-stu-id="35395-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="35395-193">String</span><span class="sxs-lookup"><span data-stu-id="35395-193">String</span></span>|<span data-ttu-id="35395-194">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="35395-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="35395-195">String</span><span class="sxs-lookup"><span data-stu-id="35395-195">String</span></span>|<span data-ttu-id="35395-196">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="35395-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35395-197">Требования</span><span class="sxs-lookup"><span data-stu-id="35395-197">Requirements</span></span>

|<span data-ttu-id="35395-198">Требование</span><span class="sxs-lookup"><span data-stu-id="35395-198">Requirement</span></span>| <span data-ttu-id="35395-199">Значение</span><span class="sxs-lookup"><span data-stu-id="35395-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="35395-200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35395-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35395-201">1.1</span><span class="sxs-lookup"><span data-stu-id="35395-201">1.1</span></span>|
|[<span data-ttu-id="35395-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35395-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35395-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35395-203">Compose or Read</span></span>|
