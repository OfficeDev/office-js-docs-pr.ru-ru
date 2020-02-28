---
title: Элемент DefaultSettings в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 824c575b39a99c6028ffd603390d2b41ee0ad7dd
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324886"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="eba28-102">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="eba28-102">DefaultSettings element</span></span>

<span data-ttu-id="eba28-103">Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="eba28-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="eba28-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="eba28-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="eba28-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="eba28-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="eba28-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="eba28-106">Contained in</span></span>

[<span data-ttu-id="eba28-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="eba28-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="eba28-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="eba28-108">Can contain</span></span>

|<span data-ttu-id="eba28-109">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="eba28-109">**Element**</span></span>|<span data-ttu-id="eba28-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="eba28-110">**Content**</span></span>|<span data-ttu-id="eba28-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="eba28-111">**Mail**</span></span>|<span data-ttu-id="eba28-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="eba28-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="eba28-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="eba28-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="eba28-114">x</span><span class="sxs-lookup"><span data-stu-id="eba28-114">x</span></span>||<span data-ttu-id="eba28-115">x</span><span class="sxs-lookup"><span data-stu-id="eba28-115">x</span></span>|
|[<span data-ttu-id="eba28-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="eba28-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="eba28-117">x</span><span class="sxs-lookup"><span data-stu-id="eba28-117">x</span></span>|||
|[<span data-ttu-id="eba28-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="eba28-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="eba28-119">x</span><span class="sxs-lookup"><span data-stu-id="eba28-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="eba28-120">Замечания</span><span class="sxs-lookup"><span data-stu-id="eba28-120">Remarks</span></span>

<span data-ttu-id="eba28-121">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к контентным надстройкам и надстройкам области задач. Для почтовых надстроек указываются расположения по умолчанию для исходных файлов и другие параметры по умолчанию в элементе [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="eba28-121">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

