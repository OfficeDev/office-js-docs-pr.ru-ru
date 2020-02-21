---
title: Пространство имен Office — набор обязательных элементов 1,1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 68363f101b4c818853cc118e39d05784c56ef3ad
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165477"
---
# <a name="office"></a><span data-ttu-id="e555b-102">Office</span><span class="sxs-lookup"><span data-stu-id="e555b-102">Office</span></span>

<span data-ttu-id="e555b-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e555b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e555b-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="e555b-105">Requirements</span></span>

|<span data-ttu-id="e555b-106">Требование</span><span class="sxs-lookup"><span data-stu-id="e555b-106">Requirement</span></span>| <span data-ttu-id="e555b-107">Значение</span><span class="sxs-lookup"><span data-stu-id="e555b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e555b-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e555b-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e555b-109">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-109">1.1</span></span>|
|[<span data-ttu-id="e555b-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e555b-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e555b-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e555b-112">Properties</span><span class="sxs-lookup"><span data-stu-id="e555b-112">Properties</span></span>

| <span data-ttu-id="e555b-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="e555b-113">Property</span></span> | <span data-ttu-id="e555b-114">Способов</span><span class="sxs-lookup"><span data-stu-id="e555b-114">Modes</span></span> | <span data-ttu-id="e555b-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e555b-115">Return type</span></span> | <span data-ttu-id="e555b-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="e555b-116">Minimum</span></span><br><span data-ttu-id="e555b-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="e555b-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e555b-118">контекст</span><span class="sxs-lookup"><span data-stu-id="e555b-118">context</span></span>](office.context.md) | <span data-ttu-id="e555b-119">Создание</span><span class="sxs-lookup"><span data-stu-id="e555b-119">Compose</span></span><br><span data-ttu-id="e555b-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-120">Read</span></span> | [<span data-ttu-id="e555b-121">Context</span><span class="sxs-lookup"><span data-stu-id="e555b-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="e555b-122">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="e555b-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="e555b-123">Enumerations</span></span>

| <span data-ttu-id="e555b-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="e555b-124">Enumeration</span></span> | <span data-ttu-id="e555b-125">Способов</span><span class="sxs-lookup"><span data-stu-id="e555b-125">Modes</span></span> | <span data-ttu-id="e555b-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e555b-126">Return type</span></span> | <span data-ttu-id="e555b-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="e555b-127">Minimum</span></span><br><span data-ttu-id="e555b-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="e555b-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e555b-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e555b-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e555b-130">Создание</span><span class="sxs-lookup"><span data-stu-id="e555b-130">Compose</span></span><br><span data-ttu-id="e555b-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-131">Read</span></span> | <span data-ttu-id="e555b-132">String</span><span class="sxs-lookup"><span data-stu-id="e555b-132">String</span></span> | [<span data-ttu-id="e555b-133">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e555b-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e555b-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e555b-135">Создание</span><span class="sxs-lookup"><span data-stu-id="e555b-135">Compose</span></span><br><span data-ttu-id="e555b-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-136">Read</span></span> | <span data-ttu-id="e555b-137">String</span><span class="sxs-lookup"><span data-stu-id="e555b-137">String</span></span> | [<span data-ttu-id="e555b-138">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e555b-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e555b-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e555b-140">Создание</span><span class="sxs-lookup"><span data-stu-id="e555b-140">Compose</span></span><br><span data-ttu-id="e555b-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-141">Read</span></span> | <span data-ttu-id="e555b-142">String</span><span class="sxs-lookup"><span data-stu-id="e555b-142">String</span></span> | [<span data-ttu-id="e555b-143">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="e555b-144">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="e555b-144">Namespaces</span></span>

<span data-ttu-id="e555b-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="e555b-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e555b-146">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="e555b-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e555b-147">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="e555b-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="e555b-148">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="e555b-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e555b-149">Тип</span><span class="sxs-lookup"><span data-stu-id="e555b-149">Type</span></span>

*   <span data-ttu-id="e555b-150">String</span><span class="sxs-lookup"><span data-stu-id="e555b-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e555b-151">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e555b-151">Properties:</span></span>

|<span data-ttu-id="e555b-152">Имя</span><span class="sxs-lookup"><span data-stu-id="e555b-152">Name</span></span>| <span data-ttu-id="e555b-153">Тип</span><span class="sxs-lookup"><span data-stu-id="e555b-153">Type</span></span>| <span data-ttu-id="e555b-154">Описание</span><span class="sxs-lookup"><span data-stu-id="e555b-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e555b-155">String</span><span class="sxs-lookup"><span data-stu-id="e555b-155">String</span></span>|<span data-ttu-id="e555b-156">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="e555b-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e555b-157">Для указания</span><span class="sxs-lookup"><span data-stu-id="e555b-157">String</span></span>|<span data-ttu-id="e555b-158">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="e555b-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e555b-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="e555b-159">Requirements</span></span>

|<span data-ttu-id="e555b-160">Требование</span><span class="sxs-lookup"><span data-stu-id="e555b-160">Requirement</span></span>| <span data-ttu-id="e555b-161">Значение</span><span class="sxs-lookup"><span data-stu-id="e555b-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="e555b-162">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e555b-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e555b-163">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-163">1.1</span></span>|
|[<span data-ttu-id="e555b-164">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e555b-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e555b-165">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e555b-166">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="e555b-166">CoercionType: String</span></span>

<span data-ttu-id="e555b-167">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e555b-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e555b-168">Тип</span><span class="sxs-lookup"><span data-stu-id="e555b-168">Type</span></span>

*   <span data-ttu-id="e555b-169">String</span><span class="sxs-lookup"><span data-stu-id="e555b-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e555b-170">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e555b-170">Properties:</span></span>

|<span data-ttu-id="e555b-171">Имя</span><span class="sxs-lookup"><span data-stu-id="e555b-171">Name</span></span>| <span data-ttu-id="e555b-172">Тип</span><span class="sxs-lookup"><span data-stu-id="e555b-172">Type</span></span>| <span data-ttu-id="e555b-173">Описание</span><span class="sxs-lookup"><span data-stu-id="e555b-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e555b-174">String</span><span class="sxs-lookup"><span data-stu-id="e555b-174">String</span></span>|<span data-ttu-id="e555b-175">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="e555b-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e555b-176">String</span><span class="sxs-lookup"><span data-stu-id="e555b-176">String</span></span>|<span data-ttu-id="e555b-177">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="e555b-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e555b-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="e555b-178">Requirements</span></span>

|<span data-ttu-id="e555b-179">Требование</span><span class="sxs-lookup"><span data-stu-id="e555b-179">Requirement</span></span>| <span data-ttu-id="e555b-180">Значение</span><span class="sxs-lookup"><span data-stu-id="e555b-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="e555b-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e555b-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e555b-182">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-182">1.1</span></span>|
|[<span data-ttu-id="e555b-183">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e555b-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e555b-184">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e555b-185">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="e555b-185">SourceProperty: String</span></span>

<span data-ttu-id="e555b-186">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e555b-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e555b-187">Тип</span><span class="sxs-lookup"><span data-stu-id="e555b-187">Type</span></span>

*   <span data-ttu-id="e555b-188">String</span><span class="sxs-lookup"><span data-stu-id="e555b-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e555b-189">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e555b-189">Properties:</span></span>

|<span data-ttu-id="e555b-190">Имя</span><span class="sxs-lookup"><span data-stu-id="e555b-190">Name</span></span>| <span data-ttu-id="e555b-191">Тип</span><span class="sxs-lookup"><span data-stu-id="e555b-191">Type</span></span>| <span data-ttu-id="e555b-192">Описание</span><span class="sxs-lookup"><span data-stu-id="e555b-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e555b-193">String</span><span class="sxs-lookup"><span data-stu-id="e555b-193">String</span></span>|<span data-ttu-id="e555b-194">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="e555b-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e555b-195">String</span><span class="sxs-lookup"><span data-stu-id="e555b-195">String</span></span>|<span data-ttu-id="e555b-196">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="e555b-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e555b-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="e555b-197">Requirements</span></span>

|<span data-ttu-id="e555b-198">Требование</span><span class="sxs-lookup"><span data-stu-id="e555b-198">Requirement</span></span>| <span data-ttu-id="e555b-199">Значение</span><span class="sxs-lookup"><span data-stu-id="e555b-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="e555b-200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e555b-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e555b-201">1.1</span><span class="sxs-lookup"><span data-stu-id="e555b-201">1.1</span></span>|
|[<span data-ttu-id="e555b-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e555b-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e555b-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e555b-203">Compose or Read</span></span>|