---
title: Элемент Icon в файле манифеста
description: Определяет элементы Image для элементов управления Button или Menu.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a17f43352b306850c853c230f6a3617eb165ca14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718092"
---
# <a name="icon-element"></a><span data-ttu-id="086e5-103">Элемент Icon</span><span class="sxs-lookup"><span data-stu-id="086e5-103">Icon element</span></span>

<span data-ttu-id="086e5-104">Определяет элементы **Image** для элементов управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="086e5-104">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="086e5-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="086e5-105">Attributes</span></span>

|  <span data-ttu-id="086e5-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="086e5-106">Attribute</span></span>  |  <span data-ttu-id="086e5-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="086e5-107">Required</span></span>  |  <span data-ttu-id="086e5-108">Описание</span><span class="sxs-lookup"><span data-stu-id="086e5-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="086e5-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="086e5-109">**xsi:type**</span></span>  |  <span data-ttu-id="086e5-110">Нет</span><span class="sxs-lookup"><span data-stu-id="086e5-110">No</span></span>  | <span data-ttu-id="086e5-p101">Тип определяемого значка. Относится только к значкам в форм-факторах мобильных устройств. Для элементов **Icon**, содержащихся в элементе [MobileFormFactor](mobileformfactor.md), этому атрибуту присвоено значение `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="086e5-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="086e5-114">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="086e5-114">Child elements</span></span>

|  <span data-ttu-id="086e5-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="086e5-115">Element</span></span> |  <span data-ttu-id="086e5-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="086e5-116">Required</span></span>  |  <span data-ttu-id="086e5-117">Описание</span><span class="sxs-lookup"><span data-stu-id="086e5-117">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="086e5-118">Image</span><span class="sxs-lookup"><span data-stu-id="086e5-118">Image</span></span>](#image)        | <span data-ttu-id="086e5-119">Да</span><span class="sxs-lookup"><span data-stu-id="086e5-119">Yes</span></span> |   <span data-ttu-id="086e5-120">атрибут resid используемого изображения</span><span class="sxs-lookup"><span data-stu-id="086e5-120">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="086e5-121">Изображение</span><span class="sxs-lookup"><span data-stu-id="086e5-121">Image</span></span>

<span data-ttu-id="086e5-122">Изображение кнопки.</span><span class="sxs-lookup"><span data-stu-id="086e5-122">An image for the button.</span></span> <span data-ttu-id="086e5-123">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **Image** **в элементе** Images элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="086e5-123">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element.</span></span> <span data-ttu-id="086e5-124">Атрибут **size** указывает размер изображения в пикселях.</span><span class="sxs-lookup"><span data-stu-id="086e5-124">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="086e5-125">Обязательными являются три размера изображения (16, 32 и 80 пикселей), а поддерживаются еще пять (20, 24, 40, 48 и 64 пикселя).|</span><span class="sxs-lookup"><span data-stu-id="086e5-125">Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="086e5-126">Дополнительные требования для форм-факторов мобильных устройств</span><span class="sxs-lookup"><span data-stu-id="086e5-126">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="086e5-p103">Когда родительский элемент **Icon** является потомком элемента [MobileFormFactor](mobileformfactor.md), минимальные требуемые размеры несколько отличаются. В манифесте должны быть указаны размеры, составляющие по крайней мере 48 x 48, 32 x 32 и 25 x 25 пикселей. Каждый указанный размер должен встречаться три раза, при этом атрибуту `scale` должно быть присвоено значение `1`, `2` или `3`.</span><span class="sxs-lookup"><span data-stu-id="086e5-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
