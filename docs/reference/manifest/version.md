---
title: Элемент Version в файле манифеста
description: Элемент Version указывает версию надстройки Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173936"
---
# <a name="version-element"></a><span data-ttu-id="9943d-103">Элемент Version</span><span class="sxs-lookup"><span data-stu-id="9943d-103">Version element</span></span>

<span data-ttu-id="9943d-104">Указывает версию надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9943d-104">Specifies the version of your Office Add-in.</span></span> <span data-ttu-id="9943d-105">Номер версии: 1, 2, 3 или 4 части (например, n, n.n, n.n.n или n.n.n).</span><span class="sxs-lookup"><span data-stu-id="9943d-105">The version number can be 1, 2, 3, or 4 parts (i.e., n, n.n, n.n.n, or n.n.n.n).</span></span>

<span data-ttu-id="9943d-106">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="9943d-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9943d-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9943d-107">Syntax</span></span>

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a><span data-ttu-id="9943d-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9943d-108">Contained in</span></span>

[<span data-ttu-id="9943d-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9943d-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="9943d-110">Замечания</span><span class="sxs-lookup"><span data-stu-id="9943d-110">Remarks</span></span>

<span data-ttu-id="9943d-111">Каждая часть номера версии может быть не более 5 цифр.</span><span class="sxs-lookup"><span data-stu-id="9943d-111">Each part of the version number can be a maximum of 5 digits.</span></span>
