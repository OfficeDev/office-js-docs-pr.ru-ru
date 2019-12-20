---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 4284f5e2367e72700d1b34bbac18b08bb0bc77f8
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814838"
---
# <a name="diagnostics"></a><span data-ttu-id="96177-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="96177-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="96177-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="96177-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="96177-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="96177-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96177-105">Требования</span><span class="sxs-lookup"><span data-stu-id="96177-105">Requirements</span></span>

|<span data-ttu-id="96177-106">Требование</span><span class="sxs-lookup"><span data-stu-id="96177-106">Requirement</span></span>| <span data-ttu-id="96177-107">Значение</span><span class="sxs-lookup"><span data-stu-id="96177-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="96177-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="96177-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96177-109">1.1</span><span class="sxs-lookup"><span data-stu-id="96177-109">1.1</span></span>|
|[<span data-ttu-id="96177-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="96177-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96177-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96177-111">ReadItem</span></span>|
|[<span data-ttu-id="96177-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="96177-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96177-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="96177-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="96177-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="96177-114">Properties</span></span>

| <span data-ttu-id="96177-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="96177-115">Property</span></span> | <span data-ttu-id="96177-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="96177-116">Minimum</span></span><br><span data-ttu-id="96177-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="96177-117">permission level</span></span> | <span data-ttu-id="96177-118">Способов</span><span class="sxs-lookup"><span data-stu-id="96177-118">Modes</span></span> | <span data-ttu-id="96177-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="96177-119">Return type</span></span> | <span data-ttu-id="96177-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="96177-120">Minimum</span></span><br><span data-ttu-id="96177-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="96177-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="96177-122">Сайту</span><span class="sxs-lookup"><span data-stu-id="96177-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5#hostname) | <span data-ttu-id="96177-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96177-123">ReadItem</span></span> | <span data-ttu-id="96177-124">Создание</span><span class="sxs-lookup"><span data-stu-id="96177-124">Compose</span></span><br><span data-ttu-id="96177-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="96177-125">Read</span></span> | <span data-ttu-id="96177-126">String</span><span class="sxs-lookup"><span data-stu-id="96177-126">String</span></span> | [<span data-ttu-id="96177-127">1.1</span><span class="sxs-lookup"><span data-stu-id="96177-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96177-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="96177-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5#hostversion) | <span data-ttu-id="96177-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96177-129">ReadItem</span></span> | <span data-ttu-id="96177-130">Создание</span><span class="sxs-lookup"><span data-stu-id="96177-130">Compose</span></span><br><span data-ttu-id="96177-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="96177-131">Read</span></span> | <span data-ttu-id="96177-132">String</span><span class="sxs-lookup"><span data-stu-id="96177-132">String</span></span> | [<span data-ttu-id="96177-133">1.1</span><span class="sxs-lookup"><span data-stu-id="96177-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96177-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="96177-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5#owaview) | <span data-ttu-id="96177-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96177-135">ReadItem</span></span> | <span data-ttu-id="96177-136">Создание</span><span class="sxs-lookup"><span data-stu-id="96177-136">Compose</span></span><br><span data-ttu-id="96177-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="96177-137">Read</span></span> | <span data-ttu-id="96177-138">String</span><span class="sxs-lookup"><span data-stu-id="96177-138">String</span></span> | [<span data-ttu-id="96177-139">1.1</span><span class="sxs-lookup"><span data-stu-id="96177-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
