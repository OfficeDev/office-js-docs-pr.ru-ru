---
title: Элемент DefaultSettings в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 199acf8be888ba51fda83d159937a74685ca48e0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450627"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="3886d-102">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="3886d-102">DefaultSettings element</span></span>

<span data-ttu-id="3886d-103">Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="3886d-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="3886d-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="3886d-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="3886d-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3886d-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="3886d-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="3886d-106">Contained in</span></span>

[<span data-ttu-id="3886d-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3886d-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="3886d-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="3886d-108">Can contain</span></span>

|<span data-ttu-id="3886d-109">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="3886d-109">**Element**</span></span>|<span data-ttu-id="3886d-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="3886d-110">**Content**</span></span>|<span data-ttu-id="3886d-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="3886d-111">**Mail**</span></span>|<span data-ttu-id="3886d-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="3886d-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3886d-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3886d-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="3886d-114">x</span><span class="sxs-lookup"><span data-stu-id="3886d-114">x</span></span>||<span data-ttu-id="3886d-115">x</span><span class="sxs-lookup"><span data-stu-id="3886d-115">x</span></span>|
|[<span data-ttu-id="3886d-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="3886d-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="3886d-117">x</span><span class="sxs-lookup"><span data-stu-id="3886d-117">x</span></span>|||
|[<span data-ttu-id="3886d-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="3886d-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="3886d-119">x</span><span class="sxs-lookup"><span data-stu-id="3886d-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="3886d-120">Замечания</span><span class="sxs-lookup"><span data-stu-id="3886d-120">Remarks</span></span>

<span data-ttu-id="3886d-121">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие стандартные параметры с помощью элемента [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="3886d-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

