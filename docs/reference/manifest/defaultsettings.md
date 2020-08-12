---
title: Элемент DefaultSettings в файле манифеста
description: Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a9711fb44390bcbda8979b8018eed1318c5579bc
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641468"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="76a3f-103">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="76a3f-103">DefaultSettings element</span></span>

<span data-ttu-id="76a3f-104">Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="76a3f-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="76a3f-105">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="76a3f-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="76a3f-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="76a3f-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="76a3f-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="76a3f-107">Contained in</span></span>

[<span data-ttu-id="76a3f-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="76a3f-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="76a3f-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="76a3f-109">Can contain</span></span>

|<span data-ttu-id="76a3f-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="76a3f-110">Element</span></span>|<span data-ttu-id="76a3f-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="76a3f-111">Content</span></span>|<span data-ttu-id="76a3f-112">Почта</span><span class="sxs-lookup"><span data-stu-id="76a3f-112">Mail</span></span>|<span data-ttu-id="76a3f-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="76a3f-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="76a3f-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="76a3f-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="76a3f-115">x</span><span class="sxs-lookup"><span data-stu-id="76a3f-115">x</span></span>||<span data-ttu-id="76a3f-116">x</span><span class="sxs-lookup"><span data-stu-id="76a3f-116">x</span></span>|
|[<span data-ttu-id="76a3f-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="76a3f-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="76a3f-118">x</span><span class="sxs-lookup"><span data-stu-id="76a3f-118">x</span></span>|||
|[<span data-ttu-id="76a3f-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="76a3f-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="76a3f-120">x</span><span class="sxs-lookup"><span data-stu-id="76a3f-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="76a3f-121">Замечания</span><span class="sxs-lookup"><span data-stu-id="76a3f-121">Remarks</span></span>

<span data-ttu-id="76a3f-122">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к контентным надстройкам и надстройкам области задач. Для почтовых надстроек указываются расположения по умолчанию для исходных файлов и другие параметры по умолчанию в элементе [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="76a3f-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>
