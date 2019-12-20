---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 6b5229c1bc300d11714f3aa2cf8fa8ff2465667c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814782"
---
# <a name="userprofile"></a><span data-ttu-id="f2f0b-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="f2f0b-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="f2f0b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="f2f0b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="f2f0b-104">Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="f2f0b-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f2f0b-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f2f0b-105">Requirements</span></span>

|<span data-ttu-id="f2f0b-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f2f0b-106">Requirement</span></span>| <span data-ttu-id="f2f0b-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f2f0b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2f0b-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f2f0b-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f2f0b-109">1.1</span><span class="sxs-lookup"><span data-stu-id="f2f0b-109">1.1</span></span>|
|[<span data-ttu-id="f2f0b-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f2f0b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f2f0b-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f2f0b-111">ReadItem</span></span>|
|[<span data-ttu-id="f2f0b-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f2f0b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f2f0b-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f2f0b-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="f2f0b-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="f2f0b-114">Properties</span></span>

| <span data-ttu-id="f2f0b-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="f2f0b-115">Property</span></span> | <span data-ttu-id="f2f0b-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f2f0b-116">Minimum</span></span><br><span data-ttu-id="f2f0b-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="f2f0b-117">permission level</span></span> | <span data-ttu-id="f2f0b-118">Способов</span><span class="sxs-lookup"><span data-stu-id="f2f0b-118">Modes</span></span> | <span data-ttu-id="f2f0b-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="f2f0b-119">Return type</span></span> | <span data-ttu-id="f2f0b-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f2f0b-120">Minimum</span></span><br><span data-ttu-id="f2f0b-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="f2f0b-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="f2f0b-122">displayName</span><span class="sxs-lookup"><span data-stu-id="f2f0b-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="f2f0b-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f2f0b-123">ReadItem</span></span> | <span data-ttu-id="f2f0b-124">Создание</span><span class="sxs-lookup"><span data-stu-id="f2f0b-124">Compose</span></span><br><span data-ttu-id="f2f0b-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="f2f0b-125">Read</span></span> | <span data-ttu-id="f2f0b-126">String</span><span class="sxs-lookup"><span data-stu-id="f2f0b-126">String</span></span> | [<span data-ttu-id="f2f0b-127">1.1</span><span class="sxs-lookup"><span data-stu-id="f2f0b-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f2f0b-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f2f0b-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="f2f0b-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f2f0b-129">ReadItem</span></span> | <span data-ttu-id="f2f0b-130">Создание</span><span class="sxs-lookup"><span data-stu-id="f2f0b-130">Compose</span></span><br><span data-ttu-id="f2f0b-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="f2f0b-131">Read</span></span> | <span data-ttu-id="f2f0b-132">String</span><span class="sxs-lookup"><span data-stu-id="f2f0b-132">String</span></span> | [<span data-ttu-id="f2f0b-133">1.1</span><span class="sxs-lookup"><span data-stu-id="f2f0b-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f2f0b-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="f2f0b-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="f2f0b-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f2f0b-135">ReadItem</span></span> | <span data-ttu-id="f2f0b-136">Создание</span><span class="sxs-lookup"><span data-stu-id="f2f0b-136">Compose</span></span><br><span data-ttu-id="f2f0b-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="f2f0b-137">Read</span></span> | <span data-ttu-id="f2f0b-138">String</span><span class="sxs-lookup"><span data-stu-id="f2f0b-138">String</span></span> | [<span data-ttu-id="f2f0b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f2f0b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
