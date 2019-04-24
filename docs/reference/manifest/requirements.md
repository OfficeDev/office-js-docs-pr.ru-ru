---
title: Элемент Requirements в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450564"
---
# <a name="requirements-element"></a><span data-ttu-id="b1273-102">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="b1273-102">Requirements element</span></span>

<span data-ttu-id="b1273-103">Указывает минимальный набор элементов API JavaScript для Office ([набор требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="b1273-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="b1273-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="b1273-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b1273-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b1273-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="b1273-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b1273-106">Contained in</span></span>

[<span data-ttu-id="b1273-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b1273-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b1273-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="b1273-108">Can contain</span></span>

|<span data-ttu-id="b1273-109">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="b1273-109">**Element**</span></span>|<span data-ttu-id="b1273-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="b1273-110">**Content**</span></span>|<span data-ttu-id="b1273-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="b1273-111">**Mail**</span></span>|<span data-ttu-id="b1273-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="b1273-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b1273-113">Sets</span><span class="sxs-lookup"><span data-stu-id="b1273-113">Sets</span></span>](sets.md)|<span data-ttu-id="b1273-114">x</span><span class="sxs-lookup"><span data-stu-id="b1273-114">x</span></span>|<span data-ttu-id="b1273-115">x</span><span class="sxs-lookup"><span data-stu-id="b1273-115">x</span></span>|<span data-ttu-id="b1273-116">x</span><span class="sxs-lookup"><span data-stu-id="b1273-116">x</span></span>|
|[<span data-ttu-id="b1273-117">Методы</span><span class="sxs-lookup"><span data-stu-id="b1273-117">Methods</span></span>](methods.md)|<span data-ttu-id="b1273-118">x</span><span class="sxs-lookup"><span data-stu-id="b1273-118">x</span></span>||<span data-ttu-id="b1273-119">x</span><span class="sxs-lookup"><span data-stu-id="b1273-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="b1273-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="b1273-120">Remarks</span></span>

<span data-ttu-id="b1273-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b1273-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

