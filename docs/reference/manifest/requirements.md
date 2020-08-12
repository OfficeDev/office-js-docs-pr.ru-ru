---
title: Элемент Requirements в файле манифеста
description: Элемент указывает минимальный набор обязательных требований и методы, необходимые надстройке Office для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c6a9a7b5923401fc2551f239b2c6cbc0d1e90755
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641321"
---
# <a name="requirements-element"></a><span data-ttu-id="e204e-103">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="e204e-103">Requirements element</span></span>

<span data-ttu-id="e204e-104">Указывает минимальный набор требований к API JavaScript для Office ([набор требований](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) и/или методов), которые должна активировать надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="e204e-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="e204e-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="e204e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e204e-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e204e-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="e204e-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e204e-107">Contained in</span></span>

[<span data-ttu-id="e204e-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e204e-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e204e-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="e204e-109">Can contain</span></span>

|<span data-ttu-id="e204e-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="e204e-110">Element</span></span>|<span data-ttu-id="e204e-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="e204e-111">Content</span></span>|<span data-ttu-id="e204e-112">Почта</span><span class="sxs-lookup"><span data-stu-id="e204e-112">Mail</span></span>|<span data-ttu-id="e204e-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="e204e-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e204e-114">Sets</span><span class="sxs-lookup"><span data-stu-id="e204e-114">Sets</span></span>](sets.md)|<span data-ttu-id="e204e-115">x</span><span class="sxs-lookup"><span data-stu-id="e204e-115">x</span></span>|<span data-ttu-id="e204e-116">x</span><span class="sxs-lookup"><span data-stu-id="e204e-116">x</span></span>|<span data-ttu-id="e204e-117">x</span><span class="sxs-lookup"><span data-stu-id="e204e-117">x</span></span>|
|[<span data-ttu-id="e204e-118">Методы</span><span class="sxs-lookup"><span data-stu-id="e204e-118">Methods</span></span>](methods.md)|<span data-ttu-id="e204e-119">x</span><span class="sxs-lookup"><span data-stu-id="e204e-119">x</span></span>||<span data-ttu-id="e204e-120">x</span><span class="sxs-lookup"><span data-stu-id="e204e-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="e204e-121">Примечания</span><span class="sxs-lookup"><span data-stu-id="e204e-121">Remarks</span></span>

<span data-ttu-id="e204e-122">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e204e-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
