---
title: Элемент DefaultSettings в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433757"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="6fa90-102">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="6fa90-102">DefaultSettings element</span></span>

<span data-ttu-id="6fa90-103">Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="6fa90-103">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="6fa90-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="6fa90-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="6fa90-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="6fa90-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="6fa90-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="6fa90-106">Contained in</span></span>

[<span data-ttu-id="6fa90-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="6fa90-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="6fa90-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="6fa90-108">Can contain</span></span>

|<span data-ttu-id="6fa90-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="6fa90-109">**Element**</span></span>|<span data-ttu-id="6fa90-110">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="6fa90-110">**Content**</span></span>|<span data-ttu-id="6fa90-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="6fa90-111">**Mail**</span></span>|<span data-ttu-id="6fa90-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="6fa90-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="6fa90-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6fa90-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="6fa90-114">x</span><span class="sxs-lookup"><span data-stu-id="6fa90-114">x</span></span>||<span data-ttu-id="6fa90-115">x</span><span class="sxs-lookup"><span data-stu-id="6fa90-115">x</span></span>|
|[<span data-ttu-id="6fa90-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="6fa90-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="6fa90-117">x</span><span class="sxs-lookup"><span data-stu-id="6fa90-117">x</span></span>|||
|[<span data-ttu-id="6fa90-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="6fa90-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="6fa90-119">x</span><span class="sxs-lookup"><span data-stu-id="6fa90-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="6fa90-120">Замечания</span><span class="sxs-lookup"><span data-stu-id="6fa90-120">Remarks</span></span>

<span data-ttu-id="6fa90-121">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие стандартные параметры с помощью элемента [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="6fa90-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

