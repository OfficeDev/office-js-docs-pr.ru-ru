---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 5ceafe65dedcb1db6c67ca28f9a1d9e05f805850
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814287"
---
# <a name="diagnostics"></a><span data-ttu-id="b6be4-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="b6be4-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="b6be4-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="b6be4-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="b6be4-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="b6be4-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b6be4-105">Требования</span><span class="sxs-lookup"><span data-stu-id="b6be4-105">Requirements</span></span>

|<span data-ttu-id="b6be4-106">Требование</span><span class="sxs-lookup"><span data-stu-id="b6be4-106">Requirement</span></span>| <span data-ttu-id="b6be4-107">Значение</span><span class="sxs-lookup"><span data-stu-id="b6be4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b6be4-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b6be4-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b6be4-109">1.1</span><span class="sxs-lookup"><span data-stu-id="b6be4-109">1.1</span></span>|
|[<span data-ttu-id="b6be4-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b6be4-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b6be4-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6be4-111">ReadItem</span></span>|
|[<span data-ttu-id="b6be4-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b6be4-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b6be4-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b6be4-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="b6be4-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="b6be4-114">Properties</span></span>

| <span data-ttu-id="b6be4-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="b6be4-115">Property</span></span> | <span data-ttu-id="b6be4-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="b6be4-116">Minimum</span></span><br><span data-ttu-id="b6be4-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="b6be4-117">permission level</span></span> | <span data-ttu-id="b6be4-118">Способов</span><span class="sxs-lookup"><span data-stu-id="b6be4-118">Modes</span></span> | <span data-ttu-id="b6be4-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="b6be4-119">Return type</span></span> | <span data-ttu-id="b6be4-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="b6be4-120">Minimum</span></span><br><span data-ttu-id="b6be4-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="b6be4-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="b6be4-122">Сайту</span><span class="sxs-lookup"><span data-stu-id="b6be4-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4#hostname) | <span data-ttu-id="b6be4-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6be4-123">ReadItem</span></span> | <span data-ttu-id="b6be4-124">Создание</span><span class="sxs-lookup"><span data-stu-id="b6be4-124">Compose</span></span><br><span data-ttu-id="b6be4-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="b6be4-125">Read</span></span> | <span data-ttu-id="b6be4-126">String</span><span class="sxs-lookup"><span data-stu-id="b6be4-126">String</span></span> | [<span data-ttu-id="b6be4-127">1.1</span><span class="sxs-lookup"><span data-stu-id="b6be4-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b6be4-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="b6be4-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4#hostversion) | <span data-ttu-id="b6be4-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6be4-129">ReadItem</span></span> | <span data-ttu-id="b6be4-130">Создание</span><span class="sxs-lookup"><span data-stu-id="b6be4-130">Compose</span></span><br><span data-ttu-id="b6be4-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="b6be4-131">Read</span></span> | <span data-ttu-id="b6be4-132">String</span><span class="sxs-lookup"><span data-stu-id="b6be4-132">String</span></span> | [<span data-ttu-id="b6be4-133">1.1</span><span class="sxs-lookup"><span data-stu-id="b6be4-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b6be4-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="b6be4-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4#owaview) | <span data-ttu-id="b6be4-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6be4-135">ReadItem</span></span> | <span data-ttu-id="b6be4-136">Создание</span><span class="sxs-lookup"><span data-stu-id="b6be4-136">Compose</span></span><br><span data-ttu-id="b6be4-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="b6be4-137">Read</span></span> | <span data-ttu-id="b6be4-138">String</span><span class="sxs-lookup"><span data-stu-id="b6be4-138">String</span></span> | [<span data-ttu-id="b6be4-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b6be4-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
