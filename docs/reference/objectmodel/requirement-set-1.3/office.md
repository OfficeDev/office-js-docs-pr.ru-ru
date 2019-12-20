---
title: Пространство имен Office — набор обязательных элементов 1.3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3c6ddc34001f4d1622bc76d9bca1fbde9425be8b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814901"
---
# <a name="office"></a><span data-ttu-id="d522a-102">Office</span><span class="sxs-lookup"><span data-stu-id="d522a-102">Office</span></span>

<span data-ttu-id="d522a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d522a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d522a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="d522a-105">Requirements</span></span>

|<span data-ttu-id="d522a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="d522a-106">Requirement</span></span>| <span data-ttu-id="d522a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="d522a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d522a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d522a-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d522a-109">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-109">1.1</span></span>|
|[<span data-ttu-id="d522a-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d522a-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d522a-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d522a-112">Properties</span><span class="sxs-lookup"><span data-stu-id="d522a-112">Properties</span></span>

| <span data-ttu-id="d522a-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="d522a-113">Property</span></span> | <span data-ttu-id="d522a-114">Способов</span><span class="sxs-lookup"><span data-stu-id="d522a-114">Modes</span></span> | <span data-ttu-id="d522a-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="d522a-115">Return type</span></span> | <span data-ttu-id="d522a-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="d522a-116">Minimum</span></span><br><span data-ttu-id="d522a-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="d522a-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d522a-118">контекст</span><span class="sxs-lookup"><span data-stu-id="d522a-118">context</span></span>](office.context.md) | <span data-ttu-id="d522a-119">Создание</span><span class="sxs-lookup"><span data-stu-id="d522a-119">Compose</span></span><br><span data-ttu-id="d522a-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-120">Read</span></span> | [<span data-ttu-id="d522a-121">Context</span><span class="sxs-lookup"><span data-stu-id="d522a-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="d522a-122">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="d522a-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="d522a-123">Enumerations</span></span>

| <span data-ttu-id="d522a-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="d522a-124">Enumeration</span></span> | <span data-ttu-id="d522a-125">Способов</span><span class="sxs-lookup"><span data-stu-id="d522a-125">Modes</span></span> | <span data-ttu-id="d522a-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="d522a-126">Return type</span></span> | <span data-ttu-id="d522a-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="d522a-127">Minimum</span></span><br><span data-ttu-id="d522a-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="d522a-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d522a-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d522a-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d522a-130">Создание</span><span class="sxs-lookup"><span data-stu-id="d522a-130">Compose</span></span><br><span data-ttu-id="d522a-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-131">Read</span></span> | <span data-ttu-id="d522a-132">String</span><span class="sxs-lookup"><span data-stu-id="d522a-132">String</span></span> | [<span data-ttu-id="d522a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d522a-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d522a-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d522a-135">Создание</span><span class="sxs-lookup"><span data-stu-id="d522a-135">Compose</span></span><br><span data-ttu-id="d522a-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-136">Read</span></span> | <span data-ttu-id="d522a-137">String</span><span class="sxs-lookup"><span data-stu-id="d522a-137">String</span></span> | [<span data-ttu-id="d522a-138">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d522a-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d522a-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d522a-140">Создание</span><span class="sxs-lookup"><span data-stu-id="d522a-140">Compose</span></span><br><span data-ttu-id="d522a-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-141">Read</span></span> | <span data-ttu-id="d522a-142">String</span><span class="sxs-lookup"><span data-stu-id="d522a-142">String</span></span> | [<span data-ttu-id="d522a-143">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="d522a-144">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d522a-144">Namespaces</span></span>

<span data-ttu-id="d522a-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="d522a-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d522a-146">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="d522a-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d522a-147">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="d522a-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="d522a-148">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="d522a-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d522a-149">Тип</span><span class="sxs-lookup"><span data-stu-id="d522a-149">Type</span></span>

*   <span data-ttu-id="d522a-150">String</span><span class="sxs-lookup"><span data-stu-id="d522a-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d522a-151">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d522a-151">Properties:</span></span>

|<span data-ttu-id="d522a-152">Имя</span><span class="sxs-lookup"><span data-stu-id="d522a-152">Name</span></span>| <span data-ttu-id="d522a-153">Тип</span><span class="sxs-lookup"><span data-stu-id="d522a-153">Type</span></span>| <span data-ttu-id="d522a-154">Описание</span><span class="sxs-lookup"><span data-stu-id="d522a-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d522a-155">String</span><span class="sxs-lookup"><span data-stu-id="d522a-155">String</span></span>|<span data-ttu-id="d522a-156">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="d522a-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d522a-157">Для указания</span><span class="sxs-lookup"><span data-stu-id="d522a-157">String</span></span>|<span data-ttu-id="d522a-158">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="d522a-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d522a-159">Требования</span><span class="sxs-lookup"><span data-stu-id="d522a-159">Requirements</span></span>

|<span data-ttu-id="d522a-160">Требование</span><span class="sxs-lookup"><span data-stu-id="d522a-160">Requirement</span></span>| <span data-ttu-id="d522a-161">Значение</span><span class="sxs-lookup"><span data-stu-id="d522a-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="d522a-162">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d522a-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d522a-163">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-163">1.1</span></span>|
|[<span data-ttu-id="d522a-164">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d522a-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d522a-165">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d522a-166">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="d522a-166">CoercionType: String</span></span>

<span data-ttu-id="d522a-167">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d522a-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d522a-168">Тип</span><span class="sxs-lookup"><span data-stu-id="d522a-168">Type</span></span>

*   <span data-ttu-id="d522a-169">String</span><span class="sxs-lookup"><span data-stu-id="d522a-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d522a-170">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d522a-170">Properties:</span></span>

|<span data-ttu-id="d522a-171">Имя</span><span class="sxs-lookup"><span data-stu-id="d522a-171">Name</span></span>| <span data-ttu-id="d522a-172">Тип</span><span class="sxs-lookup"><span data-stu-id="d522a-172">Type</span></span>| <span data-ttu-id="d522a-173">Описание</span><span class="sxs-lookup"><span data-stu-id="d522a-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d522a-174">String</span><span class="sxs-lookup"><span data-stu-id="d522a-174">String</span></span>|<span data-ttu-id="d522a-175">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d522a-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d522a-176">String</span><span class="sxs-lookup"><span data-stu-id="d522a-176">String</span></span>|<span data-ttu-id="d522a-177">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="d522a-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d522a-178">Требования</span><span class="sxs-lookup"><span data-stu-id="d522a-178">Requirements</span></span>

|<span data-ttu-id="d522a-179">Требование</span><span class="sxs-lookup"><span data-stu-id="d522a-179">Requirement</span></span>| <span data-ttu-id="d522a-180">Значение</span><span class="sxs-lookup"><span data-stu-id="d522a-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="d522a-181">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d522a-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d522a-182">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-182">1.1</span></span>|
|[<span data-ttu-id="d522a-183">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d522a-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d522a-184">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d522a-185">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="d522a-185">SourceProperty: String</span></span>

<span data-ttu-id="d522a-186">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d522a-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d522a-187">Тип</span><span class="sxs-lookup"><span data-stu-id="d522a-187">Type</span></span>

*   <span data-ttu-id="d522a-188">String</span><span class="sxs-lookup"><span data-stu-id="d522a-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d522a-189">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d522a-189">Properties:</span></span>

|<span data-ttu-id="d522a-190">Имя</span><span class="sxs-lookup"><span data-stu-id="d522a-190">Name</span></span>| <span data-ttu-id="d522a-191">Тип</span><span class="sxs-lookup"><span data-stu-id="d522a-191">Type</span></span>| <span data-ttu-id="d522a-192">Описание</span><span class="sxs-lookup"><span data-stu-id="d522a-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d522a-193">String</span><span class="sxs-lookup"><span data-stu-id="d522a-193">String</span></span>|<span data-ttu-id="d522a-194">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="d522a-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d522a-195">String</span><span class="sxs-lookup"><span data-stu-id="d522a-195">String</span></span>|<span data-ttu-id="d522a-196">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="d522a-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d522a-197">Требования</span><span class="sxs-lookup"><span data-stu-id="d522a-197">Requirements</span></span>

|<span data-ttu-id="d522a-198">Требование</span><span class="sxs-lookup"><span data-stu-id="d522a-198">Requirement</span></span>| <span data-ttu-id="d522a-199">Значение</span><span class="sxs-lookup"><span data-stu-id="d522a-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="d522a-200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d522a-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d522a-201">1.1</span><span class="sxs-lookup"><span data-stu-id="d522a-201">1.1</span></span>|
|[<span data-ttu-id="d522a-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d522a-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d522a-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d522a-203">Compose or Read</span></span>|
