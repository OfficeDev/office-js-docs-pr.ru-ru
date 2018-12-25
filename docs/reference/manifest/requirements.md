---
title: Элемент Requirements в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432706"
---
# <a name="requirements-element"></a><span data-ttu-id="d3738-102">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="d3738-102">Requirements element</span></span>

<span data-ttu-id="d3738-103">Указывает минимальный набор элементов API JavaScript для Office ([набор требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="d3738-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="d3738-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="d3738-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d3738-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="d3738-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="d3738-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="d3738-106">Contained in</span></span>

[<span data-ttu-id="d3738-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d3738-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d3738-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="d3738-108">Can contain</span></span>

|<span data-ttu-id="d3738-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="d3738-109">**Element**</span></span>|<span data-ttu-id="d3738-110">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="d3738-110">**Content**</span></span>|<span data-ttu-id="d3738-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="d3738-111">**Mail**</span></span>|<span data-ttu-id="d3738-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="d3738-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="d3738-113">Sets</span><span class="sxs-lookup"><span data-stu-id="d3738-113">Sets</span></span>](sets.md)|<span data-ttu-id="d3738-114">x</span><span class="sxs-lookup"><span data-stu-id="d3738-114">x</span></span>|<span data-ttu-id="d3738-115">x</span><span class="sxs-lookup"><span data-stu-id="d3738-115">x</span></span>|<span data-ttu-id="d3738-116">x</span><span class="sxs-lookup"><span data-stu-id="d3738-116">x</span></span>|
|[<span data-ttu-id="d3738-117">Methods</span><span class="sxs-lookup"><span data-stu-id="d3738-117">Methods</span></span>](methods.md)|<span data-ttu-id="d3738-118">x</span><span class="sxs-lookup"><span data-stu-id="d3738-118">x</span></span>||<span data-ttu-id="d3738-119">x</span><span class="sxs-lookup"><span data-stu-id="d3738-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="d3738-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="d3738-120">Remarks</span></span>

<span data-ttu-id="d3738-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d3738-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

