---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7b7b9c7facd0542335094a42a3d1f53dab1f6aef
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814322"
---
# <a name="userprofile"></a><span data-ttu-id="ec8ac-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ec8ac-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ec8ac-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ec8ac-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="ec8ac-104">Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="ec8ac-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec8ac-105">Требования</span><span class="sxs-lookup"><span data-stu-id="ec8ac-105">Requirements</span></span>

|<span data-ttu-id="ec8ac-106">Требование</span><span class="sxs-lookup"><span data-stu-id="ec8ac-106">Requirement</span></span>| <span data-ttu-id="ec8ac-107">Значение</span><span class="sxs-lookup"><span data-stu-id="ec8ac-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec8ac-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec8ac-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ec8ac-109">1.1</span><span class="sxs-lookup"><span data-stu-id="ec8ac-109">1.1</span></span>|
|[<span data-ttu-id="ec8ac-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec8ac-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec8ac-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec8ac-111">ReadItem</span></span>|
|[<span data-ttu-id="ec8ac-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec8ac-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ec8ac-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec8ac-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="ec8ac-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="ec8ac-114">Properties</span></span>

| <span data-ttu-id="ec8ac-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="ec8ac-115">Property</span></span> | <span data-ttu-id="ec8ac-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="ec8ac-116">Minimum</span></span><br><span data-ttu-id="ec8ac-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="ec8ac-117">permission level</span></span> | <span data-ttu-id="ec8ac-118">Способов</span><span class="sxs-lookup"><span data-stu-id="ec8ac-118">Modes</span></span> | <span data-ttu-id="ec8ac-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="ec8ac-119">Return type</span></span> | <span data-ttu-id="ec8ac-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="ec8ac-120">Minimum</span></span><br><span data-ttu-id="ec8ac-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="ec8ac-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="ec8ac-122">displayName</span><span class="sxs-lookup"><span data-stu-id="ec8ac-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#displayname) | <span data-ttu-id="ec8ac-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec8ac-123">ReadItem</span></span> | <span data-ttu-id="ec8ac-124">Создание</span><span class="sxs-lookup"><span data-stu-id="ec8ac-124">Compose</span></span><br><span data-ttu-id="ec8ac-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec8ac-125">Read</span></span> | <span data-ttu-id="ec8ac-126">String</span><span class="sxs-lookup"><span data-stu-id="ec8ac-126">String</span></span> | [<span data-ttu-id="ec8ac-127">1.1</span><span class="sxs-lookup"><span data-stu-id="ec8ac-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ec8ac-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ec8ac-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#emailaddress) | <span data-ttu-id="ec8ac-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec8ac-129">ReadItem</span></span> | <span data-ttu-id="ec8ac-130">Создание</span><span class="sxs-lookup"><span data-stu-id="ec8ac-130">Compose</span></span><br><span data-ttu-id="ec8ac-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec8ac-131">Read</span></span> | <span data-ttu-id="ec8ac-132">String</span><span class="sxs-lookup"><span data-stu-id="ec8ac-132">String</span></span> | [<span data-ttu-id="ec8ac-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ec8ac-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ec8ac-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="ec8ac-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#timezone) | <span data-ttu-id="ec8ac-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec8ac-135">ReadItem</span></span> | <span data-ttu-id="ec8ac-136">Создание</span><span class="sxs-lookup"><span data-stu-id="ec8ac-136">Compose</span></span><br><span data-ttu-id="ec8ac-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec8ac-137">Read</span></span> | <span data-ttu-id="ec8ac-138">String</span><span class="sxs-lookup"><span data-stu-id="ec8ac-138">String</span></span> | [<span data-ttu-id="ec8ac-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ec8ac-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
