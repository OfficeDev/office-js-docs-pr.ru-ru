---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 3872ece926cc399ed2b30d4dabaacfb741e060ab
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771405"
---
# <a name="group-element"></a><span data-ttu-id="e09b5-103">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="e09b5-103">Group element</span></span>

<span data-ttu-id="e09b5-104">Определяет группу элементов управления пользовательского интерфейса на вкладке. На настраиваемой вкладке надстройка может создать несколько групп.</span><span class="sxs-lookup"><span data-stu-id="e09b5-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="e09b5-105">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="e09b5-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="e09b5-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e09b5-106">Attributes</span></span>

|  <span data-ttu-id="e09b5-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e09b5-107">Attribute</span></span>  |  <span data-ttu-id="e09b5-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e09b5-108">Required</span></span>  |  <span data-ttu-id="e09b5-109">Описание</span><span class="sxs-lookup"><span data-stu-id="e09b5-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e09b5-110">id</span><span class="sxs-lookup"><span data-stu-id="e09b5-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="e09b5-111">Да</span><span class="sxs-lookup"><span data-stu-id="e09b5-111">Yes</span></span>  | <span data-ttu-id="e09b5-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="e09b5-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="e09b5-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="e09b5-113">id attribute</span></span>

<span data-ttu-id="e09b5-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="e09b5-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e09b5-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e09b5-118">Child elements</span></span>

|  <span data-ttu-id="e09b5-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="e09b5-119">Element</span></span> |  <span data-ttu-id="e09b5-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e09b5-120">Required</span></span>  |  <span data-ttu-id="e09b5-121">Описание</span><span class="sxs-lookup"><span data-stu-id="e09b5-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e09b5-122">Label</span><span class="sxs-lookup"><span data-stu-id="e09b5-122">Label</span></span>](#label)      | <span data-ttu-id="e09b5-123">Да</span><span class="sxs-lookup"><span data-stu-id="e09b5-123">Yes</span></span> |  <span data-ttu-id="e09b5-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="e09b5-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="e09b5-125">Icon</span><span class="sxs-lookup"><span data-stu-id="e09b5-125">Icon</span></span>](icon.md)      | <span data-ttu-id="e09b5-126">Да</span><span class="sxs-lookup"><span data-stu-id="e09b5-126">Yes</span></span> |  <span data-ttu-id="e09b5-127">Изображение для группы.</span><span class="sxs-lookup"><span data-stu-id="e09b5-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="e09b5-128">Control</span><span class="sxs-lookup"><span data-stu-id="e09b5-128">Control</span></span>](#control)    | <span data-ttu-id="e09b5-129">Нет</span><span class="sxs-lookup"><span data-stu-id="e09b5-129">No</span></span> |  <span data-ttu-id="e09b5-130">Представляет объект Control.</span><span class="sxs-lookup"><span data-stu-id="e09b5-130">Represents a Control object.</span></span> <span data-ttu-id="e09b5-131">Может иметь значение ноль или больше.</span><span class="sxs-lookup"><span data-stu-id="e09b5-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="e09b5-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="e09b5-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="e09b5-133">Нет</span><span class="sxs-lookup"><span data-stu-id="e09b5-133">No</span></span> | <span data-ttu-id="e09b5-134">Представляет один из встроенных элементов управления Office.</span><span class="sxs-lookup"><span data-stu-id="e09b5-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="e09b5-135">Может иметь значение ноль или больше.</span><span class="sxs-lookup"><span data-stu-id="e09b5-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="e09b5-136">Label</span><span class="sxs-lookup"><span data-stu-id="e09b5-136">Label</span></span>

<span data-ttu-id="e09b5-137">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="e09b5-137">Required.</span></span> <span data-ttu-id="e09b5-138">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="e09b5-138">The label of the group.</span></span> <span data-ttu-id="e09b5-139">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="e09b5-139">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="e09b5-140">Icon</span><span class="sxs-lookup"><span data-stu-id="e09b5-140">Icon</span></span>

<span data-ttu-id="e09b5-141">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="e09b5-141">Required.</span></span> <span data-ttu-id="e09b5-142">Если вкладка содержит большое количество групп и размер окна программы не задан, вместо него может отображаться указанное изображение.</span><span class="sxs-lookup"><span data-stu-id="e09b5-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="e09b5-143">Средство контроля</span><span class="sxs-lookup"><span data-stu-id="e09b5-143">Control</span></span>

<span data-ttu-id="e09b5-144">Необязательный, но если его нет, должен быть хотя бы один **OfficeControl.**</span><span class="sxs-lookup"><span data-stu-id="e09b5-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="e09b5-145">Подробные сведения о поддерживаемых типах элементов управления см. в [элементе Control.](control.md)</span><span class="sxs-lookup"><span data-stu-id="e09b5-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="e09b5-146">Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**</span><span class="sxs-lookup"><span data-stu-id="e09b5-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="officecontrol"></a><span data-ttu-id="e09b5-147">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="e09b5-147">OfficeControl</span></span>

<span data-ttu-id="e09b5-148">Необязательный, но если его нет, должен быть хотя бы один **control.**</span><span class="sxs-lookup"><span data-stu-id="e09b5-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="e09b5-149">Включаем один или несколько встроенных элементов управления Office в группу с `<OfficeControl>` элементами.</span><span class="sxs-lookup"><span data-stu-id="e09b5-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="e09b5-150">Атрибут `id` указывает ИД встроенного в Office управления.</span><span class="sxs-lookup"><span data-stu-id="e09b5-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="e09b5-151">Чтобы найти ИД элементов управления, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="e09b5-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="e09b5-152">Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**</span><span class="sxs-lookup"><span data-stu-id="e09b5-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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
