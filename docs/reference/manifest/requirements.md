---
title: Элемент Requirements в файле манифеста
description: Элемент указывает минимальный набор обязательных требований и методы, необходимые надстройке Office для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292274"
---
# <a name="requirements-element"></a><span data-ttu-id="05d7c-103">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="05d7c-103">Requirements element</span></span>

<span data-ttu-id="05d7c-104">Указывает минимальный набор требований к API JavaScript для Office ([набор требований](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) и/или методов), которые должна активировать надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="05d7c-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="05d7c-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="05d7c-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="05d7c-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="05d7c-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="05d7c-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="05d7c-107">Contained in</span></span>

[<span data-ttu-id="05d7c-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="05d7c-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="05d7c-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="05d7c-109">Can contain</span></span>

|<span data-ttu-id="05d7c-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="05d7c-110">Element</span></span>|<span data-ttu-id="05d7c-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="05d7c-111">Content</span></span>|<span data-ttu-id="05d7c-112">Почта</span><span class="sxs-lookup"><span data-stu-id="05d7c-112">Mail</span></span>|<span data-ttu-id="05d7c-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="05d7c-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="05d7c-114">Sets</span><span class="sxs-lookup"><span data-stu-id="05d7c-114">Sets</span></span>](sets.md)|<span data-ttu-id="05d7c-115">x</span><span class="sxs-lookup"><span data-stu-id="05d7c-115">x</span></span>|<span data-ttu-id="05d7c-116">x</span><span class="sxs-lookup"><span data-stu-id="05d7c-116">x</span></span>|<span data-ttu-id="05d7c-117">x</span><span class="sxs-lookup"><span data-stu-id="05d7c-117">x</span></span>|
|[<span data-ttu-id="05d7c-118">Методы</span><span class="sxs-lookup"><span data-stu-id="05d7c-118">Methods</span></span>](methods.md)|<span data-ttu-id="05d7c-119">x</span><span class="sxs-lookup"><span data-stu-id="05d7c-119">x</span></span>||<span data-ttu-id="05d7c-120">x</span><span class="sxs-lookup"><span data-stu-id="05d7c-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="05d7c-121">Примечания</span><span class="sxs-lookup"><span data-stu-id="05d7c-121">Remarks</span></span>

<span data-ttu-id="05d7c-122">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="05d7c-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
