---
title: Элемент Requirements в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3c4cb81ebd6a38ea311e8fcacfa6d5fcd3b26f68
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325250"
---
# <a name="requirements-element"></a><span data-ttu-id="629ae-102">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="629ae-102">Requirements element</span></span>

<span data-ttu-id="629ae-103">Указывает минимальный набор требований к API JavaScript для Office ([набор требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), которые должна активировать надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="629ae-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="629ae-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="629ae-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="629ae-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="629ae-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="629ae-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="629ae-106">Contained in</span></span>

[<span data-ttu-id="629ae-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="629ae-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="629ae-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="629ae-108">Can contain</span></span>

|<span data-ttu-id="629ae-109">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="629ae-109">**Element**</span></span>|<span data-ttu-id="629ae-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="629ae-110">**Content**</span></span>|<span data-ttu-id="629ae-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="629ae-111">**Mail**</span></span>|<span data-ttu-id="629ae-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="629ae-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="629ae-113">Sets</span><span class="sxs-lookup"><span data-stu-id="629ae-113">Sets</span></span>](sets.md)|<span data-ttu-id="629ae-114">x</span><span class="sxs-lookup"><span data-stu-id="629ae-114">x</span></span>|<span data-ttu-id="629ae-115">x</span><span class="sxs-lookup"><span data-stu-id="629ae-115">x</span></span>|<span data-ttu-id="629ae-116">x</span><span class="sxs-lookup"><span data-stu-id="629ae-116">x</span></span>|
|[<span data-ttu-id="629ae-117">Методы</span><span class="sxs-lookup"><span data-stu-id="629ae-117">Methods</span></span>](methods.md)|<span data-ttu-id="629ae-118">x</span><span class="sxs-lookup"><span data-stu-id="629ae-118">x</span></span>||<span data-ttu-id="629ae-119">x</span><span class="sxs-lookup"><span data-stu-id="629ae-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="629ae-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="629ae-120">Remarks</span></span>

<span data-ttu-id="629ae-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="629ae-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

