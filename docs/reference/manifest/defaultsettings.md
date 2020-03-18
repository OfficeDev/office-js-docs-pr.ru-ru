---
title: Элемент DefaultSettings в файле манифеста
description: Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b97f692a1fd39e4b1f55080f6ed77e623be0000c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718372"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="cfd4a-103">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="cfd4a-103">DefaultSettings element</span></span>

<span data-ttu-id="cfd4a-104">Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="cfd4a-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="cfd4a-105">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="cfd4a-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="cfd4a-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="cfd4a-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="cfd4a-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="cfd4a-107">Contained in</span></span>

[<span data-ttu-id="cfd4a-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cfd4a-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="cfd4a-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="cfd4a-109">Can contain</span></span>

|<span data-ttu-id="cfd4a-110">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="cfd4a-110">**Element**</span></span>|<span data-ttu-id="cfd4a-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="cfd4a-111">**Content**</span></span>|<span data-ttu-id="cfd4a-112">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="cfd4a-112">**Mail**</span></span>|<span data-ttu-id="cfd4a-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="cfd4a-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="cfd4a-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cfd4a-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="cfd4a-115">x</span><span class="sxs-lookup"><span data-stu-id="cfd4a-115">x</span></span>||<span data-ttu-id="cfd4a-116">x</span><span class="sxs-lookup"><span data-stu-id="cfd4a-116">x</span></span>|
|[<span data-ttu-id="cfd4a-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="cfd4a-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="cfd4a-118">x</span><span class="sxs-lookup"><span data-stu-id="cfd4a-118">x</span></span>|||
|[<span data-ttu-id="cfd4a-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="cfd4a-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="cfd4a-120">x</span><span class="sxs-lookup"><span data-stu-id="cfd4a-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="cfd4a-121">Замечания</span><span class="sxs-lookup"><span data-stu-id="cfd4a-121">Remarks</span></span>

<span data-ttu-id="cfd4a-122">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к контентным надстройкам и надстройкам области задач. Для почтовых надстроек указываются расположения по умолчанию для исходных файлов и другие параметры по умолчанию в элементе [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="cfd4a-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

