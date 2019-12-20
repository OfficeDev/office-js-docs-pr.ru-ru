---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 1bf24eb39329be0139957cc6e0f8629fb9f3b166
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815020"
---
# <a name="userprofile"></a><span data-ttu-id="0115c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="0115c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="0115c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="0115c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="0115c-104">Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="0115c-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0115c-105">Требования</span><span class="sxs-lookup"><span data-stu-id="0115c-105">Requirements</span></span>

|<span data-ttu-id="0115c-106">Требование</span><span class="sxs-lookup"><span data-stu-id="0115c-106">Requirement</span></span>| <span data-ttu-id="0115c-107">Значение</span><span class="sxs-lookup"><span data-stu-id="0115c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0115c-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0115c-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0115c-109">1.1</span><span class="sxs-lookup"><span data-stu-id="0115c-109">1.1</span></span>|
|[<span data-ttu-id="0115c-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0115c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0115c-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0115c-111">ReadItem</span></span>|
|[<span data-ttu-id="0115c-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0115c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0115c-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0115c-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="0115c-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="0115c-114">Properties</span></span>

| <span data-ttu-id="0115c-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="0115c-115">Property</span></span> | <span data-ttu-id="0115c-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="0115c-116">Minimum</span></span><br><span data-ttu-id="0115c-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="0115c-117">permission level</span></span> | <span data-ttu-id="0115c-118">Способов</span><span class="sxs-lookup"><span data-stu-id="0115c-118">Modes</span></span> | <span data-ttu-id="0115c-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="0115c-119">Return type</span></span> | <span data-ttu-id="0115c-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="0115c-120">Minimum</span></span><br><span data-ttu-id="0115c-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="0115c-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="0115c-122">displayName</span><span class="sxs-lookup"><span data-stu-id="0115c-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1#displayname) | <span data-ttu-id="0115c-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0115c-123">ReadItem</span></span> | <span data-ttu-id="0115c-124">Создание</span><span class="sxs-lookup"><span data-stu-id="0115c-124">Compose</span></span><br><span data-ttu-id="0115c-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="0115c-125">Read</span></span> | <span data-ttu-id="0115c-126">String</span><span class="sxs-lookup"><span data-stu-id="0115c-126">String</span></span> | [<span data-ttu-id="0115c-127">1.1</span><span class="sxs-lookup"><span data-stu-id="0115c-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0115c-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="0115c-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1#emailaddress) | <span data-ttu-id="0115c-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0115c-129">ReadItem</span></span> | <span data-ttu-id="0115c-130">Создание</span><span class="sxs-lookup"><span data-stu-id="0115c-130">Compose</span></span><br><span data-ttu-id="0115c-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="0115c-131">Read</span></span> | <span data-ttu-id="0115c-132">String</span><span class="sxs-lookup"><span data-stu-id="0115c-132">String</span></span> | [<span data-ttu-id="0115c-133">1.1</span><span class="sxs-lookup"><span data-stu-id="0115c-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0115c-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="0115c-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1#timezone) | <span data-ttu-id="0115c-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0115c-135">ReadItem</span></span> | <span data-ttu-id="0115c-136">Создание</span><span class="sxs-lookup"><span data-stu-id="0115c-136">Compose</span></span><br><span data-ttu-id="0115c-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="0115c-137">Read</span></span> | <span data-ttu-id="0115c-138">String</span><span class="sxs-lookup"><span data-stu-id="0115c-138">String</span></span> | [<span data-ttu-id="0115c-139">1.1</span><span class="sxs-lookup"><span data-stu-id="0115c-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
