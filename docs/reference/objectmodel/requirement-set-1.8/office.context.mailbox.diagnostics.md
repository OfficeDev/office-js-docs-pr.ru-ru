---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2c5ea33ffd8bc560288935f7ee65ebb93aadf1aa
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814252"
---
# <a name="diagnostics"></a><span data-ttu-id="3897a-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="3897a-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="3897a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="3897a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="3897a-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="3897a-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3897a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="3897a-105">Requirements</span></span>

|<span data-ttu-id="3897a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="3897a-106">Requirement</span></span>| <span data-ttu-id="3897a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="3897a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3897a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3897a-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3897a-109">1.1</span><span class="sxs-lookup"><span data-stu-id="3897a-109">1.1</span></span>|
|[<span data-ttu-id="3897a-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3897a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3897a-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3897a-111">ReadItem</span></span>|
|[<span data-ttu-id="3897a-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3897a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3897a-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3897a-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="3897a-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="3897a-114">Properties</span></span>

| <span data-ttu-id="3897a-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="3897a-115">Property</span></span> | <span data-ttu-id="3897a-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="3897a-116">Minimum</span></span><br><span data-ttu-id="3897a-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="3897a-117">permission level</span></span> | <span data-ttu-id="3897a-118">Способов</span><span class="sxs-lookup"><span data-stu-id="3897a-118">Modes</span></span> | <span data-ttu-id="3897a-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="3897a-119">Return type</span></span> | <span data-ttu-id="3897a-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="3897a-120">Minimum</span></span><br><span data-ttu-id="3897a-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="3897a-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="3897a-122">Сайту</span><span class="sxs-lookup"><span data-stu-id="3897a-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#hostname) | <span data-ttu-id="3897a-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3897a-123">ReadItem</span></span> | <span data-ttu-id="3897a-124">Создание</span><span class="sxs-lookup"><span data-stu-id="3897a-124">Compose</span></span><br><span data-ttu-id="3897a-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="3897a-125">Read</span></span> | <span data-ttu-id="3897a-126">String</span><span class="sxs-lookup"><span data-stu-id="3897a-126">String</span></span> | [<span data-ttu-id="3897a-127">1.1</span><span class="sxs-lookup"><span data-stu-id="3897a-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3897a-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="3897a-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#hostversion) | <span data-ttu-id="3897a-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3897a-129">ReadItem</span></span> | <span data-ttu-id="3897a-130">Создание</span><span class="sxs-lookup"><span data-stu-id="3897a-130">Compose</span></span><br><span data-ttu-id="3897a-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="3897a-131">Read</span></span> | <span data-ttu-id="3897a-132">String</span><span class="sxs-lookup"><span data-stu-id="3897a-132">String</span></span> | [<span data-ttu-id="3897a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="3897a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3897a-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="3897a-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#owaview) | <span data-ttu-id="3897a-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3897a-135">ReadItem</span></span> | <span data-ttu-id="3897a-136">Создание</span><span class="sxs-lookup"><span data-stu-id="3897a-136">Compose</span></span><br><span data-ttu-id="3897a-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="3897a-137">Read</span></span> | <span data-ttu-id="3897a-138">String</span><span class="sxs-lookup"><span data-stu-id="3897a-138">String</span></span> | [<span data-ttu-id="3897a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3897a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
