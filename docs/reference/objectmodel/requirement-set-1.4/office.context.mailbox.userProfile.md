---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950994"
---
# <a name="userprofile"></a><span data-ttu-id="a6e08-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a6e08-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a6e08-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a6e08-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="a6e08-104">Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6e08-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6e08-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6e08-105">Requirements</span></span>

|<span data-ttu-id="a6e08-106">Требование</span><span class="sxs-lookup"><span data-stu-id="a6e08-106">Requirement</span></span>| <span data-ttu-id="a6e08-107">Значение</span><span class="sxs-lookup"><span data-stu-id="a6e08-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6e08-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6e08-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6e08-109">1.1</span><span class="sxs-lookup"><span data-stu-id="a6e08-109">1.1</span></span>|
|[<span data-ttu-id="a6e08-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a6e08-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6e08-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6e08-111">ReadItem</span></span>|
|[<span data-ttu-id="a6e08-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6e08-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6e08-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6e08-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="a6e08-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="a6e08-114">Properties</span></span>

| <span data-ttu-id="a6e08-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="a6e08-115">Property</span></span> | <span data-ttu-id="a6e08-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="a6e08-116">Minimum</span></span><br><span data-ttu-id="a6e08-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="a6e08-117">permission level</span></span> | <span data-ttu-id="a6e08-118">Способов</span><span class="sxs-lookup"><span data-stu-id="a6e08-118">Modes</span></span> | <span data-ttu-id="a6e08-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="a6e08-119">Return type</span></span> | <span data-ttu-id="a6e08-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="a6e08-120">Minimum</span></span><br><span data-ttu-id="a6e08-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="a6e08-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="a6e08-122">displayName</span><span class="sxs-lookup"><span data-stu-id="a6e08-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="a6e08-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6e08-123">ReadItem</span></span> | <span data-ttu-id="a6e08-124">Создание</span><span class="sxs-lookup"><span data-stu-id="a6e08-124">Compose</span></span><br><span data-ttu-id="a6e08-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6e08-125">Read</span></span> | <span data-ttu-id="a6e08-126">String</span><span class="sxs-lookup"><span data-stu-id="a6e08-126">String</span></span> | [<span data-ttu-id="a6e08-127">1.1</span><span class="sxs-lookup"><span data-stu-id="a6e08-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6e08-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a6e08-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="a6e08-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6e08-129">ReadItem</span></span> | <span data-ttu-id="a6e08-130">Создание</span><span class="sxs-lookup"><span data-stu-id="a6e08-130">Compose</span></span><br><span data-ttu-id="a6e08-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6e08-131">Read</span></span> | <span data-ttu-id="a6e08-132">String</span><span class="sxs-lookup"><span data-stu-id="a6e08-132">String</span></span> | [<span data-ttu-id="a6e08-133">1.1</span><span class="sxs-lookup"><span data-stu-id="a6e08-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6e08-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="a6e08-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="a6e08-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6e08-135">ReadItem</span></span> | <span data-ttu-id="a6e08-136">Создание</span><span class="sxs-lookup"><span data-stu-id="a6e08-136">Compose</span></span><br><span data-ttu-id="a6e08-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6e08-137">Read</span></span> | <span data-ttu-id="a6e08-138">String</span><span class="sxs-lookup"><span data-stu-id="a6e08-138">String</span></span> | [<span data-ttu-id="a6e08-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a6e08-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
