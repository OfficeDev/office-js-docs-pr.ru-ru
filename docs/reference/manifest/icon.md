---
title: Элемент Icon в файле манифеста
description: Определяет элементы Image для элементов управления Button или Menu.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 1adfbcd154091fcae49966f0c1f7d0b9cc968ed3
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604626"
---
# <a name="icon-element"></a><span data-ttu-id="cdcbd-103">Элемент Icon</span><span class="sxs-lookup"><span data-stu-id="cdcbd-103">Icon element</span></span>

<span data-ttu-id="cdcbd-104">Определяет элементы **Image** для элементов управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="cdcbd-104">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="cdcbd-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="cdcbd-105">Attributes</span></span>

|  <span data-ttu-id="cdcbd-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="cdcbd-106">Attribute</span></span>  |  <span data-ttu-id="cdcbd-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cdcbd-107">Required</span></span>  |  <span data-ttu-id="cdcbd-108">Описание</span><span class="sxs-lookup"><span data-stu-id="cdcbd-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="cdcbd-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="cdcbd-109">**xsi:type**</span></span>  |  <span data-ttu-id="cdcbd-110">Нет</span><span class="sxs-lookup"><span data-stu-id="cdcbd-110">No</span></span>  | <span data-ttu-id="cdcbd-p101">Тип определяемого значка. Относится только к значкам в форм-факторах мобильных устройств. Для элементов **Icon**, содержащихся в элементе [MobileFormFactor](mobileformfactor.md), этому атрибуту присвоено значение `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="cdcbd-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="cdcbd-114">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="cdcbd-114">Child elements</span></span>

|  <span data-ttu-id="cdcbd-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="cdcbd-115">Element</span></span> |  <span data-ttu-id="cdcbd-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cdcbd-116">Required</span></span>  |  <span data-ttu-id="cdcbd-117">Описание</span><span class="sxs-lookup"><span data-stu-id="cdcbd-117">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cdcbd-118">Image</span><span class="sxs-lookup"><span data-stu-id="cdcbd-118">Image</span></span>](#image)        | <span data-ttu-id="cdcbd-119">Да</span><span class="sxs-lookup"><span data-stu-id="cdcbd-119">Yes</span></span> |   <span data-ttu-id="cdcbd-120">атрибут resid используемого изображения</span><span class="sxs-lookup"><span data-stu-id="cdcbd-120">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="cdcbd-121">Изображение</span><span class="sxs-lookup"><span data-stu-id="cdcbd-121">Image</span></span>

<span data-ttu-id="cdcbd-122">Изображение кнопки.</span><span class="sxs-lookup"><span data-stu-id="cdcbd-122">An image for the button.</span></span> <span data-ttu-id="cdcbd-123">Атрибут **resid** может быть не более 32 символов и должен быть задатки значению атрибута **id** элемента **Image** в элементе **Images** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="cdcbd-123">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element.</span></span> <span data-ttu-id="cdcbd-124">Атрибут **size** указывает размер изображения в пикселях.</span><span class="sxs-lookup"><span data-stu-id="cdcbd-124">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="cdcbd-125">Необходимо использовать три размера изображения (16, 32 и 80 пикселей), а поддерживается еще пять размеров (20, 24, 40, 48 и 64 пикселя).</span><span class="sxs-lookup"><span data-stu-id="cdcbd-125">Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> <span data-ttu-id="cdcbd-126">Если это изображение является представителем значка надстройки, см. в приложении [Create effective listings in AppSource и Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) для размера и других требований.</span><span class="sxs-lookup"><span data-stu-id="cdcbd-126">If this image is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="cdcbd-127">Дополнительные требования для форм-факторов мобильных устройств</span><span class="sxs-lookup"><span data-stu-id="cdcbd-127">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="cdcbd-p103">Когда родительский элемент **Icon** является потомком элемента [MobileFormFactor](mobileformfactor.md), минимальные требуемые размеры несколько отличаются. В манифесте должны быть указаны размеры, составляющие по крайней мере 48 x 48, 32 x 32 и 25 x 25 пикселей. Каждый указанный размер должен встречаться три раза, при этом атрибуту `scale` должно быть присвоено значение `1`, `2` или `3`.</span><span class="sxs-lookup"><span data-stu-id="cdcbd-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

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
