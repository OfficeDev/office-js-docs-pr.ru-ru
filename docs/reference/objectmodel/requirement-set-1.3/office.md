---
title: Пространство имен Office — набор обязательных элементов 1.3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b6a8c581670692ed48c4dcc2a7e1f86196b5bce7
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165435"
---
# <a name="office"></a><span data-ttu-id="b5858-102">Office</span><span class="sxs-lookup"><span data-stu-id="b5858-102">Office</span></span>

<span data-ttu-id="b5858-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b5858-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b5858-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="b5858-105">Requirements</span></span>

|<span data-ttu-id="b5858-106">Требование</span><span class="sxs-lookup"><span data-stu-id="b5858-106">Requirement</span></span>| <span data-ttu-id="b5858-107">Значение</span><span class="sxs-lookup"><span data-stu-id="b5858-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5858-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b5858-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5858-109">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-109">1.1</span></span>|
|[<span data-ttu-id="b5858-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b5858-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5858-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b5858-112">Properties</span><span class="sxs-lookup"><span data-stu-id="b5858-112">Properties</span></span>

| <span data-ttu-id="b5858-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="b5858-113">Property</span></span> | <span data-ttu-id="b5858-114">Способов</span><span class="sxs-lookup"><span data-stu-id="b5858-114">Modes</span></span> | <span data-ttu-id="b5858-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="b5858-115">Return type</span></span> | <span data-ttu-id="b5858-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="b5858-116">Minimum</span></span><br><span data-ttu-id="b5858-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="b5858-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b5858-118">контекст</span><span class="sxs-lookup"><span data-stu-id="b5858-118">context</span></span>](office.context.md) | <span data-ttu-id="b5858-119">Создание</span><span class="sxs-lookup"><span data-stu-id="b5858-119">Compose</span></span><br><span data-ttu-id="b5858-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-120">Read</span></span> | [<span data-ttu-id="b5858-121">Context</span><span class="sxs-lookup"><span data-stu-id="b5858-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="b5858-122">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="b5858-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="b5858-123">Enumerations</span></span>

| <span data-ttu-id="b5858-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="b5858-124">Enumeration</span></span> | <span data-ttu-id="b5858-125">Способов</span><span class="sxs-lookup"><span data-stu-id="b5858-125">Modes</span></span> | <span data-ttu-id="b5858-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="b5858-126">Return type</span></span> | <span data-ttu-id="b5858-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="b5858-127">Minimum</span></span><br><span data-ttu-id="b5858-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="b5858-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b5858-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b5858-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b5858-130">Создание</span><span class="sxs-lookup"><span data-stu-id="b5858-130">Compose</span></span><br><span data-ttu-id="b5858-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-131">Read</span></span> | <span data-ttu-id="b5858-132">String</span><span class="sxs-lookup"><span data-stu-id="b5858-132">String</span></span> | [<span data-ttu-id="b5858-133">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b5858-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b5858-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b5858-135">Создание</span><span class="sxs-lookup"><span data-stu-id="b5858-135">Compose</span></span><br><span data-ttu-id="b5858-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-136">Read</span></span> | <span data-ttu-id="b5858-137">String</span><span class="sxs-lookup"><span data-stu-id="b5858-137">String</span></span> | [<span data-ttu-id="b5858-138">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b5858-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b5858-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b5858-140">Создание</span><span class="sxs-lookup"><span data-stu-id="b5858-140">Compose</span></span><br><span data-ttu-id="b5858-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-141">Read</span></span> | <span data-ttu-id="b5858-142">String</span><span class="sxs-lookup"><span data-stu-id="b5858-142">String</span></span> | [<span data-ttu-id="b5858-143">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="b5858-144">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="b5858-144">Namespaces</span></span>

<span data-ttu-id="b5858-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="b5858-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="b5858-146">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="b5858-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b5858-147">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="b5858-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="b5858-148">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="b5858-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b5858-149">Тип</span><span class="sxs-lookup"><span data-stu-id="b5858-149">Type</span></span>

*   <span data-ttu-id="b5858-150">String</span><span class="sxs-lookup"><span data-stu-id="b5858-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b5858-151">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b5858-151">Properties:</span></span>

|<span data-ttu-id="b5858-152">Имя</span><span class="sxs-lookup"><span data-stu-id="b5858-152">Name</span></span>| <span data-ttu-id="b5858-153">Тип</span><span class="sxs-lookup"><span data-stu-id="b5858-153">Type</span></span>| <span data-ttu-id="b5858-154">Описание</span><span class="sxs-lookup"><span data-stu-id="b5858-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b5858-155">String</span><span class="sxs-lookup"><span data-stu-id="b5858-155">String</span></span>|<span data-ttu-id="b5858-156">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="b5858-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b5858-157">Для указания</span><span class="sxs-lookup"><span data-stu-id="b5858-157">String</span></span>|<span data-ttu-id="b5858-158">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="b5858-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b5858-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="b5858-159">Requirements</span></span>

|<span data-ttu-id="b5858-160">Требование</span><span class="sxs-lookup"><span data-stu-id="b5858-160">Requirement</span></span>| <span data-ttu-id="b5858-161">Значение</span><span class="sxs-lookup"><span data-stu-id="b5858-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5858-162">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b5858-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5858-163">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-163">1.1</span></span>|
|[<span data-ttu-id="b5858-164">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b5858-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5858-165">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="b5858-166">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="b5858-166">CoercionType: String</span></span>

<span data-ttu-id="b5858-167">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b5858-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b5858-168">Тип</span><span class="sxs-lookup"><span data-stu-id="b5858-168">Type</span></span>

*   <span data-ttu-id="b5858-169">String</span><span class="sxs-lookup"><span data-stu-id="b5858-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b5858-170">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b5858-170">Properties:</span></span>

|<span data-ttu-id="b5858-171">Имя</span><span class="sxs-lookup"><span data-stu-id="b5858-171">Name</span></span>| <span data-ttu-id="b5858-172">Тип</span><span class="sxs-lookup"><span data-stu-id="b5858-172">Type</span></span>| <span data-ttu-id="b5858-173">Описание</span><span class="sxs-lookup"><span data-stu-id="b5858-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b5858-174">String</span><span class="sxs-lookup"><span data-stu-id="b5858-174">String</span></span>|<span data-ttu-id="b5858-175">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="b5858-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b5858-176">String</span><span class="sxs-lookup"><span data-stu-id="b5858-176">String</span></span>|<span data-ttu-id="b5858-177">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="b5858-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b5858-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="b5858-178">Requirements</span></span>

|<span data-ttu-id="b5858-179">Требование</span><span class="sxs-lookup"><span data-stu-id="b5858-179">Requirement</span></span>| <span data-ttu-id="b5858-180">Значение</span><span class="sxs-lookup"><span data-stu-id="b5858-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5858-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b5858-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5858-182">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-182">1.1</span></span>|
|[<span data-ttu-id="b5858-183">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b5858-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5858-184">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="b5858-185">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="b5858-185">SourceProperty: String</span></span>

<span data-ttu-id="b5858-186">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b5858-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b5858-187">Тип</span><span class="sxs-lookup"><span data-stu-id="b5858-187">Type</span></span>

*   <span data-ttu-id="b5858-188">String</span><span class="sxs-lookup"><span data-stu-id="b5858-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b5858-189">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b5858-189">Properties:</span></span>

|<span data-ttu-id="b5858-190">Имя</span><span class="sxs-lookup"><span data-stu-id="b5858-190">Name</span></span>| <span data-ttu-id="b5858-191">Тип</span><span class="sxs-lookup"><span data-stu-id="b5858-191">Type</span></span>| <span data-ttu-id="b5858-192">Описание</span><span class="sxs-lookup"><span data-stu-id="b5858-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b5858-193">String</span><span class="sxs-lookup"><span data-stu-id="b5858-193">String</span></span>|<span data-ttu-id="b5858-194">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="b5858-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b5858-195">String</span><span class="sxs-lookup"><span data-stu-id="b5858-195">String</span></span>|<span data-ttu-id="b5858-196">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="b5858-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b5858-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="b5858-197">Requirements</span></span>

|<span data-ttu-id="b5858-198">Требование</span><span class="sxs-lookup"><span data-stu-id="b5858-198">Requirement</span></span>| <span data-ttu-id="b5858-199">Значение</span><span class="sxs-lookup"><span data-stu-id="b5858-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5858-200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b5858-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5858-201">1.1</span><span class="sxs-lookup"><span data-stu-id="b5858-201">1.1</span></span>|
|[<span data-ttu-id="b5858-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b5858-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5858-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b5858-203">Compose or Read</span></span>|
