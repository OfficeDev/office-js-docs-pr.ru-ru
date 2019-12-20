---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 4d63dfe1b32de2ac7fe55f324f938b85a865ec02
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814915"
---
# <a name="userprofile"></a><span data-ttu-id="3e383-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="3e383-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="3e383-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="3e383-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="3e383-104">Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3e383-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3e383-105">Требования</span><span class="sxs-lookup"><span data-stu-id="3e383-105">Requirements</span></span>

|<span data-ttu-id="3e383-106">Требование</span><span class="sxs-lookup"><span data-stu-id="3e383-106">Requirement</span></span>| <span data-ttu-id="3e383-107">Значение</span><span class="sxs-lookup"><span data-stu-id="3e383-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3e383-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3e383-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3e383-109">1.1</span><span class="sxs-lookup"><span data-stu-id="3e383-109">1.1</span></span>|
|[<span data-ttu-id="3e383-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3e383-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3e383-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3e383-111">ReadItem</span></span>|
|[<span data-ttu-id="3e383-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3e383-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3e383-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3e383-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="3e383-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="3e383-114">Properties</span></span>

| <span data-ttu-id="3e383-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="3e383-115">Property</span></span> | <span data-ttu-id="3e383-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="3e383-116">Minimum</span></span><br><span data-ttu-id="3e383-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="3e383-117">permission level</span></span> | <span data-ttu-id="3e383-118">Способов</span><span class="sxs-lookup"><span data-stu-id="3e383-118">Modes</span></span> | <span data-ttu-id="3e383-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="3e383-119">Return type</span></span> | <span data-ttu-id="3e383-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="3e383-120">Minimum</span></span><br><span data-ttu-id="3e383-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="3e383-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="3e383-122">displayName</span><span class="sxs-lookup"><span data-stu-id="3e383-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#displayname) | <span data-ttu-id="3e383-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3e383-123">ReadItem</span></span> | <span data-ttu-id="3e383-124">Создание</span><span class="sxs-lookup"><span data-stu-id="3e383-124">Compose</span></span><br><span data-ttu-id="3e383-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="3e383-125">Read</span></span> | <span data-ttu-id="3e383-126">String</span><span class="sxs-lookup"><span data-stu-id="3e383-126">String</span></span> | [<span data-ttu-id="3e383-127">1.1</span><span class="sxs-lookup"><span data-stu-id="3e383-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3e383-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="3e383-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#emailaddress) | <span data-ttu-id="3e383-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3e383-129">ReadItem</span></span> | <span data-ttu-id="3e383-130">Создание</span><span class="sxs-lookup"><span data-stu-id="3e383-130">Compose</span></span><br><span data-ttu-id="3e383-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="3e383-131">Read</span></span> | <span data-ttu-id="3e383-132">String</span><span class="sxs-lookup"><span data-stu-id="3e383-132">String</span></span> | [<span data-ttu-id="3e383-133">1.1</span><span class="sxs-lookup"><span data-stu-id="3e383-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3e383-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="3e383-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#timezone) | <span data-ttu-id="3e383-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3e383-135">ReadItem</span></span> | <span data-ttu-id="3e383-136">Создание</span><span class="sxs-lookup"><span data-stu-id="3e383-136">Compose</span></span><br><span data-ttu-id="3e383-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="3e383-137">Read</span></span> | <span data-ttu-id="3e383-138">String</span><span class="sxs-lookup"><span data-stu-id="3e383-138">String</span></span> | [<span data-ttu-id="3e383-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3e383-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
