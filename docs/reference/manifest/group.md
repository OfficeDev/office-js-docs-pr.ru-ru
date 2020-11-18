---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087947"
---
# <a name="group-element"></a><span data-ttu-id="738d3-103">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="738d3-103">Group element</span></span>

<span data-ttu-id="738d3-104">Определяет группу элементов управления пользовательского интерфейса на вкладке. На пользовательских вкладках надстройка может создавать несколько групп.</span><span class="sxs-lookup"><span data-stu-id="738d3-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="738d3-105">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="738d3-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="738d3-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="738d3-106">Attributes</span></span>

|  <span data-ttu-id="738d3-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="738d3-107">Attribute</span></span>  |  <span data-ttu-id="738d3-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="738d3-108">Required</span></span>  |  <span data-ttu-id="738d3-109">Описание</span><span class="sxs-lookup"><span data-stu-id="738d3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="738d3-110">id</span><span class="sxs-lookup"><span data-stu-id="738d3-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="738d3-111">Да</span><span class="sxs-lookup"><span data-stu-id="738d3-111">Yes</span></span>  | <span data-ttu-id="738d3-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="738d3-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="738d3-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="738d3-113">id attribute</span></span>

<span data-ttu-id="738d3-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="738d3-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="738d3-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="738d3-118">Child elements</span></span>

|  <span data-ttu-id="738d3-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="738d3-119">Element</span></span> |  <span data-ttu-id="738d3-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="738d3-120">Required</span></span>  |  <span data-ttu-id="738d3-121">Описание</span><span class="sxs-lookup"><span data-stu-id="738d3-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="738d3-122">Label</span><span class="sxs-lookup"><span data-stu-id="738d3-122">Label</span></span>](#label)      | <span data-ttu-id="738d3-123">Да</span><span class="sxs-lookup"><span data-stu-id="738d3-123">Yes</span></span> |  <span data-ttu-id="738d3-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="738d3-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="738d3-125">Icon</span><span class="sxs-lookup"><span data-stu-id="738d3-125">Icon</span></span>](icon.md)      | <span data-ttu-id="738d3-126">Да</span><span class="sxs-lookup"><span data-stu-id="738d3-126">Yes</span></span> |  <span data-ttu-id="738d3-127">Изображение для группы.</span><span class="sxs-lookup"><span data-stu-id="738d3-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="738d3-128">Control</span><span class="sxs-lookup"><span data-stu-id="738d3-128">Control</span></span>](#control)    | <span data-ttu-id="738d3-129">Нет</span><span class="sxs-lookup"><span data-stu-id="738d3-129">No</span></span> |  <span data-ttu-id="738d3-130">Представляет объект элемента управления.</span><span class="sxs-lookup"><span data-stu-id="738d3-130">Represents a Control object.</span></span> <span data-ttu-id="738d3-131">Может быть нулевым или более.</span><span class="sxs-lookup"><span data-stu-id="738d3-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="738d3-132">оффицеконтрол</span><span class="sxs-lookup"><span data-stu-id="738d3-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="738d3-133">Нет</span><span class="sxs-lookup"><span data-stu-id="738d3-133">No</span></span> | <span data-ttu-id="738d3-134">Представляет один из встроенных элементов управления Office.</span><span class="sxs-lookup"><span data-stu-id="738d3-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="738d3-135">Может быть нулевым или более.</span><span class="sxs-lookup"><span data-stu-id="738d3-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="738d3-136">Label</span><span class="sxs-lookup"><span data-stu-id="738d3-136">Label</span></span>

<span data-ttu-id="738d3-137">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="738d3-137">Required.</span></span> <span data-ttu-id="738d3-138">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="738d3-138">The label of the group.</span></span> <span data-ttu-id="738d3-139">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="738d3-139">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="738d3-140">Icon</span><span class="sxs-lookup"><span data-stu-id="738d3-140">Icon</span></span>

<span data-ttu-id="738d3-141">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="738d3-141">Required.</span></span> <span data-ttu-id="738d3-142">Если вкладка содержит большое количество групп и изменяется размер окна программы, вместо этого может отображаться указанное изображение.</span><span class="sxs-lookup"><span data-stu-id="738d3-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="738d3-143">Элемент управления</span><span class="sxs-lookup"><span data-stu-id="738d3-143">Control</span></span>

<span data-ttu-id="738d3-144">Необязательный параметр, но если он отсутствует, должен существовать хотя бы один **оффицеконтрол**.</span><span class="sxs-lookup"><span data-stu-id="738d3-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="738d3-145">Дополнительные сведения о поддерживаемых типах элементов управления приведены в элементе [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="738d3-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="738d3-146">Порядок **управления** и **оффицеконтрол** в манифесте являются взаимозаменяемыми и могут быть интерминглед, если существует несколько элементов, но они должны находиться под элементом **Icon** .</span><span class="sxs-lookup"><span data-stu-id="738d3-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a><span data-ttu-id="738d3-147">оффицеконтрол</span><span class="sxs-lookup"><span data-stu-id="738d3-147">OfficeControl</span></span>

<span data-ttu-id="738d3-148">Необязательный параметр, но если он отсутствует, то должен существовать по крайней мере один **элемент управления**.</span><span class="sxs-lookup"><span data-stu-id="738d3-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="738d3-149">Включите один или несколько встроенных элементов управления Office в группу с `<OfficeControl>` элементами.</span><span class="sxs-lookup"><span data-stu-id="738d3-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="738d3-150">`id`Атрибут ЗАДАЕТ Идентификатор встроенного элемента управления Office.</span><span class="sxs-lookup"><span data-stu-id="738d3-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="738d3-151">Чтобы найти идентификатор элемента управления, ознакомьтесь со статьей [Поиск идентификаторов элементов управления и групп элементов управления](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="738d3-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="738d3-152">Порядок **управления** и **оффицеконтрол** в манифесте являются взаимозаменяемыми и могут быть интерминглед, если существует несколько элементов, но они должны находиться под элементом **Icon** .</span><span class="sxs-lookup"><span data-stu-id="738d3-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```
