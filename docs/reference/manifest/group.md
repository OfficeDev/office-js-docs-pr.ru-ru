# <a name="group-element"></a><span data-ttu-id="5f520-101">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="5f520-101">Group element</span></span>

<span data-ttu-id="5f520-p101">Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="5f520-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="5f520-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f520-105">Attributes</span></span>

|  <span data-ttu-id="5f520-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="5f520-106">Attribute</span></span>  |  <span data-ttu-id="5f520-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5f520-107">Required</span></span>  |  <span data-ttu-id="5f520-108">Описание</span><span class="sxs-lookup"><span data-stu-id="5f520-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5f520-109">id</span><span class="sxs-lookup"><span data-stu-id="5f520-109">id</span></span>](#id-attribute)  |  <span data-ttu-id="5f520-110">Да</span><span class="sxs-lookup"><span data-stu-id="5f520-110">Yes</span></span>  | <span data-ttu-id="5f520-111">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="5f520-111">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="5f520-112">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="5f520-112">id attribute</span></span>

<span data-ttu-id="5f520-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="5f520-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5f520-117">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="5f520-117">Child elements</span></span>
|  <span data-ttu-id="5f520-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="5f520-118">Element</span></span> |  <span data-ttu-id="5f520-119">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5f520-119">Required</span></span>  |  <span data-ttu-id="5f520-120">Описание</span><span class="sxs-lookup"><span data-stu-id="5f520-120">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5f520-121">Label</span><span class="sxs-lookup"><span data-stu-id="5f520-121">Label</span></span>](#label)      | <span data-ttu-id="5f520-122">Да</span><span class="sxs-lookup"><span data-stu-id="5f520-122">Yes</span></span> |  <span data-ttu-id="5f520-123">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="5f520-123">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="5f520-124">Control</span><span class="sxs-lookup"><span data-stu-id="5f520-124">Control</span></span>](#control)    | <span data-ttu-id="5f520-125">Да</span><span class="sxs-lookup"><span data-stu-id="5f520-125">Yes</span></span> |  <span data-ttu-id="5f520-126">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="5f520-126">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="5f520-127">Label</span><span class="sxs-lookup"><span data-stu-id="5f520-127">Label</span></span> 

<span data-ttu-id="5f520-p103">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="5f520-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="5f520-131">Control</span><span class="sxs-lookup"><span data-stu-id="5f520-131">Control</span></span>
<span data-ttu-id="5f520-132">Для группы требуется по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="5f520-132">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```