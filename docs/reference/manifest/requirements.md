---
title: Элемент Requirements в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870368"
---
# <a name="requirements-element"></a><span data-ttu-id="9723c-102">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="9723c-102">Requirements element</span></span>

<span data-ttu-id="9723c-103">Указывает минимальный набор элементов API JavaScript для Office ([набор требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9723c-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="9723c-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="9723c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9723c-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9723c-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="9723c-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9723c-106">Contained in</span></span>

[<span data-ttu-id="9723c-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9723c-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="9723c-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="9723c-108">Can contain</span></span>

|<span data-ttu-id="9723c-109">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="9723c-109">**Element**</span></span>|<span data-ttu-id="9723c-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="9723c-110">**Content**</span></span>|<span data-ttu-id="9723c-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="9723c-111">**Mail**</span></span>|<span data-ttu-id="9723c-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="9723c-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="9723c-113">Sets</span><span class="sxs-lookup"><span data-stu-id="9723c-113">Sets</span></span>](sets.md)|<span data-ttu-id="9723c-114">x</span><span class="sxs-lookup"><span data-stu-id="9723c-114">x</span></span>|<span data-ttu-id="9723c-115">x</span><span class="sxs-lookup"><span data-stu-id="9723c-115">x</span></span>|<span data-ttu-id="9723c-116">x</span><span class="sxs-lookup"><span data-stu-id="9723c-116">x</span></span>|
|[<span data-ttu-id="9723c-117">Методы</span><span class="sxs-lookup"><span data-stu-id="9723c-117">Methods</span></span>](methods.md)|<span data-ttu-id="9723c-118">x</span><span class="sxs-lookup"><span data-stu-id="9723c-118">x</span></span>||<span data-ttu-id="9723c-119">x</span><span class="sxs-lookup"><span data-stu-id="9723c-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="9723c-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="9723c-120">Remarks</span></span>

<span data-ttu-id="9723c-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="9723c-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

