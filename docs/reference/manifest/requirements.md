---
title: Элемент Requirements в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 43c66118b9129c4c8ae395254ea82ef1cbcbaab1
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596461"
---
# <a name="requirements-element"></a><span data-ttu-id="9d07e-102">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="9d07e-102">Requirements element</span></span>

<span data-ttu-id="9d07e-103">Указывает минимальный набор требований к API JavaScript для Office ([набор требований](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) и/или методов), которые должна активировать надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="9d07e-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="9d07e-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="9d07e-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9d07e-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9d07e-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="9d07e-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9d07e-106">Contained in</span></span>

[<span data-ttu-id="9d07e-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9d07e-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="9d07e-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="9d07e-108">Can contain</span></span>

|<span data-ttu-id="9d07e-109">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="9d07e-109">**Element**</span></span>|<span data-ttu-id="9d07e-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="9d07e-110">**Content**</span></span>|<span data-ttu-id="9d07e-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="9d07e-111">**Mail**</span></span>|<span data-ttu-id="9d07e-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="9d07e-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="9d07e-113">Sets</span><span class="sxs-lookup"><span data-stu-id="9d07e-113">Sets</span></span>](sets.md)|<span data-ttu-id="9d07e-114">x</span><span class="sxs-lookup"><span data-stu-id="9d07e-114">x</span></span>|<span data-ttu-id="9d07e-115">x</span><span class="sxs-lookup"><span data-stu-id="9d07e-115">x</span></span>|<span data-ttu-id="9d07e-116">x</span><span class="sxs-lookup"><span data-stu-id="9d07e-116">x</span></span>|
|[<span data-ttu-id="9d07e-117">Методы</span><span class="sxs-lookup"><span data-stu-id="9d07e-117">Methods</span></span>](methods.md)|<span data-ttu-id="9d07e-118">x</span><span class="sxs-lookup"><span data-stu-id="9d07e-118">x</span></span>||<span data-ttu-id="9d07e-119">x</span><span class="sxs-lookup"><span data-stu-id="9d07e-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="9d07e-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="9d07e-120">Remarks</span></span>

<span data-ttu-id="9d07e-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9d07e-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
