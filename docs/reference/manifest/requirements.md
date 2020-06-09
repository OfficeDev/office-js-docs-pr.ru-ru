---
title: Элемент Requirements в файле манифеста
description: Элемент указывает минимальный набор обязательных требований и методы, необходимые надстройке Office для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611717"
---
# <a name="requirements-element"></a><span data-ttu-id="3ceb7-103">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="3ceb7-103">Requirements element</span></span>

<span data-ttu-id="3ceb7-104">Указывает минимальный набор требований к API JavaScript для Office ([набор требований](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) и/или методов), которые должна активировать надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="3ceb7-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="3ceb7-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="3ceb7-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3ceb7-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3ceb7-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="3ceb7-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="3ceb7-107">Contained in</span></span>

[<span data-ttu-id="3ceb7-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3ceb7-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="3ceb7-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="3ceb7-109">Can contain</span></span>

|<span data-ttu-id="3ceb7-110">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="3ceb7-110">**Element**</span></span>|<span data-ttu-id="3ceb7-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="3ceb7-111">**Content**</span></span>|<span data-ttu-id="3ceb7-112">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="3ceb7-112">**Mail**</span></span>|<span data-ttu-id="3ceb7-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="3ceb7-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3ceb7-114">Sets</span><span class="sxs-lookup"><span data-stu-id="3ceb7-114">Sets</span></span>](sets.md)|<span data-ttu-id="3ceb7-115">x</span><span class="sxs-lookup"><span data-stu-id="3ceb7-115">x</span></span>|<span data-ttu-id="3ceb7-116">x</span><span class="sxs-lookup"><span data-stu-id="3ceb7-116">x</span></span>|<span data-ttu-id="3ceb7-117">x</span><span class="sxs-lookup"><span data-stu-id="3ceb7-117">x</span></span>|
|[<span data-ttu-id="3ceb7-118">Методы</span><span class="sxs-lookup"><span data-stu-id="3ceb7-118">Methods</span></span>](methods.md)|<span data-ttu-id="3ceb7-119">x</span><span class="sxs-lookup"><span data-stu-id="3ceb7-119">x</span></span>||<span data-ttu-id="3ceb7-120">x</span><span class="sxs-lookup"><span data-stu-id="3ceb7-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="3ceb7-121">Примечания</span><span class="sxs-lookup"><span data-stu-id="3ceb7-121">Remarks</span></span>

<span data-ttu-id="3ceb7-122">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3ceb7-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
