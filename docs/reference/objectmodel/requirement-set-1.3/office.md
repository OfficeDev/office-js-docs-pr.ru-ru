---
title: Пространство имен Office — набор обязательных элементов 1.3
description: Office членов пространства имен, доступных для Outlook надстройки с помощью API почтовых ящиков, установленного 1.3.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f4aecf016e259141fd8adb2683864d4c36bdaf4b
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592005"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="de7ad-103">Office (набор требований к почтовым ящикам 1.3)</span><span class="sxs-lookup"><span data-stu-id="de7ad-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="de7ad-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="de7ad-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="de7ad-106">Требования</span><span class="sxs-lookup"><span data-stu-id="de7ad-106">Requirements</span></span>

|<span data-ttu-id="de7ad-107">Требование</span><span class="sxs-lookup"><span data-stu-id="de7ad-107">Requirement</span></span>| <span data-ttu-id="de7ad-108">Значение</span><span class="sxs-lookup"><span data-stu-id="de7ad-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="de7ad-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="de7ad-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="de7ad-110">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-110">1.1</span></span>|
|[<span data-ttu-id="de7ad-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="de7ad-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="de7ad-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="de7ad-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="de7ad-113">Properties</span></span>

| <span data-ttu-id="de7ad-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="de7ad-114">Property</span></span> | <span data-ttu-id="de7ad-115">Режимы</span><span class="sxs-lookup"><span data-stu-id="de7ad-115">Modes</span></span> | <span data-ttu-id="de7ad-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="de7ad-116">Return type</span></span> | <span data-ttu-id="de7ad-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="de7ad-117">Minimum</span></span><br><span data-ttu-id="de7ad-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="de7ad-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="de7ad-119">контекст</span><span class="sxs-lookup"><span data-stu-id="de7ad-119">context</span></span>](office.context.md) | <span data-ttu-id="de7ad-120">Создание</span><span class="sxs-lookup"><span data-stu-id="de7ad-120">Compose</span></span><br><span data-ttu-id="de7ad-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-121">Read</span></span> | [<span data-ttu-id="de7ad-122">Context</span><span class="sxs-lookup"><span data-stu-id="de7ad-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="de7ad-123">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="de7ad-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="de7ad-124">Enumerations</span></span>

| <span data-ttu-id="de7ad-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="de7ad-125">Enumeration</span></span> | <span data-ttu-id="de7ad-126">Режимы</span><span class="sxs-lookup"><span data-stu-id="de7ad-126">Modes</span></span> | <span data-ttu-id="de7ad-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="de7ad-127">Return type</span></span> | <span data-ttu-id="de7ad-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="de7ad-128">Minimum</span></span><br><span data-ttu-id="de7ad-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="de7ad-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="de7ad-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="de7ad-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="de7ad-131">Создание</span><span class="sxs-lookup"><span data-stu-id="de7ad-131">Compose</span></span><br><span data-ttu-id="de7ad-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-132">Read</span></span> | <span data-ttu-id="de7ad-133">Строка</span><span class="sxs-lookup"><span data-stu-id="de7ad-133">String</span></span> | [<span data-ttu-id="de7ad-134">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="de7ad-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="de7ad-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="de7ad-136">Создание</span><span class="sxs-lookup"><span data-stu-id="de7ad-136">Compose</span></span><br><span data-ttu-id="de7ad-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-137">Read</span></span> | <span data-ttu-id="de7ad-138">Строка</span><span class="sxs-lookup"><span data-stu-id="de7ad-138">String</span></span> | [<span data-ttu-id="de7ad-139">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="de7ad-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="de7ad-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="de7ad-141">Создание</span><span class="sxs-lookup"><span data-stu-id="de7ad-141">Compose</span></span><br><span data-ttu-id="de7ad-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-142">Read</span></span> | <span data-ttu-id="de7ad-143">Строка</span><span class="sxs-lookup"><span data-stu-id="de7ad-143">String</span></span> | [<span data-ttu-id="de7ad-144">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="de7ad-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="de7ad-145">Namespaces</span></span>

<span data-ttu-id="de7ad-146">[MailboxEnums:](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true)включает ряд Outlook определенных списков, например , , `ItemType` `EntityType` , `AttachmentType` , , , `RecipientType` и `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="de7ad-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="de7ad-147">Сведения о переумериях</span><span class="sxs-lookup"><span data-stu-id="de7ad-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="de7ad-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="de7ad-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="de7ad-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="de7ad-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="de7ad-150">Тип</span><span class="sxs-lookup"><span data-stu-id="de7ad-150">Type</span></span>

*   <span data-ttu-id="de7ad-151">String</span><span class="sxs-lookup"><span data-stu-id="de7ad-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="de7ad-152">Свойства</span><span class="sxs-lookup"><span data-stu-id="de7ad-152">Properties</span></span>

|<span data-ttu-id="de7ad-153">Имя</span><span class="sxs-lookup"><span data-stu-id="de7ad-153">Name</span></span>| <span data-ttu-id="de7ad-154">Тип</span><span class="sxs-lookup"><span data-stu-id="de7ad-154">Type</span></span>| <span data-ttu-id="de7ad-155">Описание</span><span class="sxs-lookup"><span data-stu-id="de7ad-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="de7ad-156">Строка</span><span class="sxs-lookup"><span data-stu-id="de7ad-156">String</span></span>|<span data-ttu-id="de7ad-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="de7ad-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="de7ad-158">String</span><span class="sxs-lookup"><span data-stu-id="de7ad-158">String</span></span>|<span data-ttu-id="de7ad-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="de7ad-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de7ad-160">Требования</span><span class="sxs-lookup"><span data-stu-id="de7ad-160">Requirements</span></span>

|<span data-ttu-id="de7ad-161">Требование</span><span class="sxs-lookup"><span data-stu-id="de7ad-161">Requirement</span></span>| <span data-ttu-id="de7ad-162">Значение</span><span class="sxs-lookup"><span data-stu-id="de7ad-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="de7ad-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="de7ad-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="de7ad-164">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-164">1.1</span></span>|
|[<span data-ttu-id="de7ad-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="de7ad-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="de7ad-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="de7ad-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="de7ad-167">CoercionType: String</span></span>

<span data-ttu-id="de7ad-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="de7ad-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="de7ad-169">Тип</span><span class="sxs-lookup"><span data-stu-id="de7ad-169">Type</span></span>

*   <span data-ttu-id="de7ad-170">String</span><span class="sxs-lookup"><span data-stu-id="de7ad-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="de7ad-171">Свойства</span><span class="sxs-lookup"><span data-stu-id="de7ad-171">Properties</span></span>

|<span data-ttu-id="de7ad-172">Имя</span><span class="sxs-lookup"><span data-stu-id="de7ad-172">Name</span></span>| <span data-ttu-id="de7ad-173">Тип</span><span class="sxs-lookup"><span data-stu-id="de7ad-173">Type</span></span>| <span data-ttu-id="de7ad-174">Описание</span><span class="sxs-lookup"><span data-stu-id="de7ad-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="de7ad-175">Строка</span><span class="sxs-lookup"><span data-stu-id="de7ad-175">String</span></span>|<span data-ttu-id="de7ad-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="de7ad-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="de7ad-177">String</span><span class="sxs-lookup"><span data-stu-id="de7ad-177">String</span></span>|<span data-ttu-id="de7ad-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="de7ad-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de7ad-179">Требования</span><span class="sxs-lookup"><span data-stu-id="de7ad-179">Requirements</span></span>

|<span data-ttu-id="de7ad-180">Требование</span><span class="sxs-lookup"><span data-stu-id="de7ad-180">Requirement</span></span>| <span data-ttu-id="de7ad-181">Значение</span><span class="sxs-lookup"><span data-stu-id="de7ad-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="de7ad-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="de7ad-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="de7ad-183">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-183">1.1</span></span>|
|[<span data-ttu-id="de7ad-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="de7ad-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="de7ad-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="de7ad-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="de7ad-186">SourceProperty: String</span></span>

<span data-ttu-id="de7ad-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="de7ad-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="de7ad-188">Тип</span><span class="sxs-lookup"><span data-stu-id="de7ad-188">Type</span></span>

*   <span data-ttu-id="de7ad-189">String</span><span class="sxs-lookup"><span data-stu-id="de7ad-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="de7ad-190">Свойства</span><span class="sxs-lookup"><span data-stu-id="de7ad-190">Properties</span></span>

|<span data-ttu-id="de7ad-191">Имя</span><span class="sxs-lookup"><span data-stu-id="de7ad-191">Name</span></span>| <span data-ttu-id="de7ad-192">Тип</span><span class="sxs-lookup"><span data-stu-id="de7ad-192">Type</span></span>| <span data-ttu-id="de7ad-193">Описание</span><span class="sxs-lookup"><span data-stu-id="de7ad-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="de7ad-194">Строка</span><span class="sxs-lookup"><span data-stu-id="de7ad-194">String</span></span>|<span data-ttu-id="de7ad-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="de7ad-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="de7ad-196">String</span><span class="sxs-lookup"><span data-stu-id="de7ad-196">String</span></span>|<span data-ttu-id="de7ad-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="de7ad-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de7ad-198">Требования</span><span class="sxs-lookup"><span data-stu-id="de7ad-198">Requirements</span></span>

|<span data-ttu-id="de7ad-199">Требование</span><span class="sxs-lookup"><span data-stu-id="de7ad-199">Requirement</span></span>| <span data-ttu-id="de7ad-200">Значение</span><span class="sxs-lookup"><span data-stu-id="de7ad-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="de7ad-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="de7ad-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="de7ad-202">1.1</span><span class="sxs-lookup"><span data-stu-id="de7ad-202">1.1</span></span>|
|[<span data-ttu-id="de7ad-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="de7ad-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="de7ad-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="de7ad-204">Compose or Read</span></span>|
