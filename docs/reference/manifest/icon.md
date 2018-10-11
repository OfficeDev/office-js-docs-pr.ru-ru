# <a name="icon-element"></a><span data-ttu-id="c2507-101">Элемент Icon</span><span class="sxs-lookup"><span data-stu-id="c2507-101">Icon element</span></span>

<span data-ttu-id="c2507-102">Определяет элементы **Image** для элементов управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="c2507-102">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="c2507-103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c2507-103">Attributes</span></span>

|  <span data-ttu-id="c2507-104">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c2507-104">Attribute</span></span>  |  <span data-ttu-id="c2507-105">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c2507-105">Required</span></span>  |  <span data-ttu-id="c2507-106">Описание</span><span class="sxs-lookup"><span data-stu-id="c2507-106">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c2507-107">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c2507-107">**xsi:type**</span></span>  |  <span data-ttu-id="c2507-108">Нет</span><span class="sxs-lookup"><span data-stu-id="c2507-108">No</span></span>  | <span data-ttu-id="c2507-p101">Тип определяемого значка. Относится только к значкам в форм-факторах мобильных устройств. Для элементов **Icon**, содержащихся в элементе [MobileFormFactor](mobileformfactor.md), этому атрибуту присвоено значение `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="c2507-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="c2507-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c2507-112">Child elements</span></span>

|  <span data-ttu-id="c2507-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="c2507-113">Element</span></span> |  <span data-ttu-id="c2507-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c2507-114">Required</span></span>  |  <span data-ttu-id="c2507-115">Описание</span><span class="sxs-lookup"><span data-stu-id="c2507-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c2507-116">Image</span><span class="sxs-lookup"><span data-stu-id="c2507-116">Image</span></span>](#image)        | <span data-ttu-id="c2507-117">Да</span><span class="sxs-lookup"><span data-stu-id="c2507-117">Yes</span></span> |   <span data-ttu-id="c2507-118">атрибут resid используемого изображения</span><span class="sxs-lookup"><span data-stu-id="c2507-118">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="c2507-119">Image</span><span class="sxs-lookup"><span data-stu-id="c2507-119">Image</span></span>

<span data-ttu-id="c2507-p102">Изображение для кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **Image** в элементе **Images** в элементе [Resources](resources.md). Атрибут **size** указывает размер изображения в пикселях. Обязательными являются три размера изображения (16, 32 и 80 пикселей), а поддерживаются еще пять (20, 24, 40, 48 и 64 пикселя).</span><span class="sxs-lookup"><span data-stu-id="c2507-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="c2507-124">Дополнительные требования для форм-факторов мобильных устройств</span><span class="sxs-lookup"><span data-stu-id="c2507-124">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="c2507-p103">Когда родительский элемент **Icon** является потомком элемента [MobileFormFactor](mobileformfactor.md), минимальные требуемые размеры несколько отличаются. В манифесте должны быть указаны по крайней мере размеры в 48, 32 и 25 пикселей. Каждый указанный размер должен встречаться три раза, при этом атрибуту `scale` должно быть присвоено значение `1`, `2` или `3`.</span><span class="sxs-lookup"><span data-stu-id="c2507-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

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