---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173964"
---
# <a name="group-element"></a><span data-ttu-id="64983-103">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="64983-103">Group element</span></span>

<span data-ttu-id="64983-104">Определяет группу элементов управления пользовательского интерфейса на вкладке. На настраиваемой вкладке надстройка может создать несколько групп.</span><span class="sxs-lookup"><span data-stu-id="64983-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="64983-105">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="64983-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="64983-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="64983-106">Attributes</span></span>

|  <span data-ttu-id="64983-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="64983-107">Attribute</span></span>  |  <span data-ttu-id="64983-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="64983-108">Required</span></span>  |  <span data-ttu-id="64983-109">Описание</span><span class="sxs-lookup"><span data-stu-id="64983-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="64983-110">id</span><span class="sxs-lookup"><span data-stu-id="64983-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="64983-111">Да</span><span class="sxs-lookup"><span data-stu-id="64983-111">Yes</span></span>  | <span data-ttu-id="64983-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="64983-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="64983-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="64983-113">id attribute</span></span>

<span data-ttu-id="64983-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="64983-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="64983-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="64983-118">Child elements</span></span>

|  <span data-ttu-id="64983-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="64983-119">Element</span></span> |  <span data-ttu-id="64983-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="64983-120">Required</span></span>  |  <span data-ttu-id="64983-121">Описание</span><span class="sxs-lookup"><span data-stu-id="64983-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="64983-122">Label</span><span class="sxs-lookup"><span data-stu-id="64983-122">Label</span></span>](#label)      | <span data-ttu-id="64983-123">Да</span><span class="sxs-lookup"><span data-stu-id="64983-123">Yes</span></span> |  <span data-ttu-id="64983-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="64983-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="64983-125">Icon</span><span class="sxs-lookup"><span data-stu-id="64983-125">Icon</span></span>](icon.md)      | <span data-ttu-id="64983-126">Да</span><span class="sxs-lookup"><span data-stu-id="64983-126">Yes</span></span> |  <span data-ttu-id="64983-127">Изображение для группы.</span><span class="sxs-lookup"><span data-stu-id="64983-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="64983-128">Control</span><span class="sxs-lookup"><span data-stu-id="64983-128">Control</span></span>](#control)    | <span data-ttu-id="64983-129">Нет</span><span class="sxs-lookup"><span data-stu-id="64983-129">No</span></span> |  <span data-ttu-id="64983-130">Представляет объект Control.</span><span class="sxs-lookup"><span data-stu-id="64983-130">Represents a Control object.</span></span> <span data-ttu-id="64983-131">Может быть ноль или больше.</span><span class="sxs-lookup"><span data-stu-id="64983-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="64983-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="64983-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="64983-133">Нет</span><span class="sxs-lookup"><span data-stu-id="64983-133">No</span></span> | <span data-ttu-id="64983-134">Представляет один из встроенных элементов управления Office.</span><span class="sxs-lookup"><span data-stu-id="64983-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="64983-135">Может быть ноль или больше.</span><span class="sxs-lookup"><span data-stu-id="64983-135">Can be zero or more.</span></span> |
|  [<span data-ttu-id="64983-136">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="64983-136">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="64983-137">Нет</span><span class="sxs-lookup"><span data-stu-id="64983-137">No</span></span> |  <span data-ttu-id="64983-138">Указывает, должна ли группа отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="64983-138">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span>  |

### <a name="label"></a><span data-ttu-id="64983-139">Label</span><span class="sxs-lookup"><span data-stu-id="64983-139">Label</span></span>

<span data-ttu-id="64983-140">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="64983-140">Required.</span></span> <span data-ttu-id="64983-141">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="64983-141">The label of the group.</span></span> <span data-ttu-id="64983-142">Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="64983-142">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="64983-143">Icon</span><span class="sxs-lookup"><span data-stu-id="64983-143">Icon</span></span>

<span data-ttu-id="64983-144">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="64983-144">Required.</span></span> <span data-ttu-id="64983-145">Если вкладка содержит большое количество групп и размер окна программы будет меняться, вместо него может отображаться указанное изображение.</span><span class="sxs-lookup"><span data-stu-id="64983-145">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="64983-146">Элемент управления</span><span class="sxs-lookup"><span data-stu-id="64983-146">Control</span></span>

<span data-ttu-id="64983-147">Необязательный, но если его нет, должен быть хотя бы один **OfficeControl.**</span><span class="sxs-lookup"><span data-stu-id="64983-147">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="64983-148">Подробные сведения о поддерживаемых типах элементов управления см. в [элементе Control.](control.md)</span><span class="sxs-lookup"><span data-stu-id="64983-148">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="64983-149">Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**</span><span class="sxs-lookup"><span data-stu-id="64983-149">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

### <a name="officecontrol"></a><span data-ttu-id="64983-150">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="64983-150">OfficeControl</span></span>

<span data-ttu-id="64983-151">Необязательный, но если его нет, должен быть хотя бы один **control.**</span><span class="sxs-lookup"><span data-stu-id="64983-151">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="64983-152">Включаем один или несколько встроенных элементов управления Office в группу с `<OfficeControl>` элементами.</span><span class="sxs-lookup"><span data-stu-id="64983-152">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="64983-153">Атрибут `id` указывает ИД встроенного в Office управления.</span><span class="sxs-lookup"><span data-stu-id="64983-153">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="64983-154">Чтобы найти ИД элементов управления, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="64983-154">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="64983-155">Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**</span><span class="sxs-lookup"><span data-stu-id="64983-155">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="64983-156">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="64983-156">OverriddenByRibbonApi</span></span>

<span data-ttu-id="64983-157">Необязательный (boolean).</span><span class="sxs-lookup"><span data-stu-id="64983-157">Optional (boolean).</span></span> <span data-ttu-id="64983-158">Указывает, будет  ли группа скрыта в сочетаниях приложений и платформ, которые поддерживают API, устанавливая настраиваемую контекстную вкладку на ленту во время работы.</span><span class="sxs-lookup"><span data-stu-id="64983-158">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="64983-159">Значение по умолчанию (если его нет) `false` — .</span><span class="sxs-lookup"><span data-stu-id="64983-159">The default value, if not present, is `false`.</span></span> <span data-ttu-id="64983-160">Если используется, **OverriddenByRibbonApi** должен быть первым *в* **группе.**</span><span class="sxs-lookup"><span data-stu-id="64983-160">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="64983-161">Дополнительные сведения [см. в подразделе OverriddenByRibbonApi.](overriddenbyribbonapi.md)</span><span class="sxs-lookup"><span data-stu-id="64983-161">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
