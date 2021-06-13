---
title: Групповой элемент в файле манифеста
description: Определяет группу элементов управления пользовательским интерфейсом на вкладке.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 89ed16f7996ab06bd21e1ebaa71c959b11af2029
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893514"
---
# <a name="group-element"></a><span data-ttu-id="4b62c-103">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="4b62c-103">Group element</span></span>

<span data-ttu-id="4b62c-104">Определяет группу элементов управления пользовательским интерфейсом на вкладке. На настраиваемой вкладке надстройка может создавать несколько групп.</span><span class="sxs-lookup"><span data-stu-id="4b62c-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="4b62c-105">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="4b62c-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b62c-106">Attributes</span></span>

|  <span data-ttu-id="4b62c-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="4b62c-107">Attribute</span></span>  |  <span data-ttu-id="4b62c-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4b62c-108">Required</span></span>  |  <span data-ttu-id="4b62c-109">Описание</span><span class="sxs-lookup"><span data-stu-id="4b62c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4b62c-110">id</span><span class="sxs-lookup"><span data-stu-id="4b62c-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="4b62c-111">Да</span><span class="sxs-lookup"><span data-stu-id="4b62c-111">Yes</span></span>  | <span data-ttu-id="4b62c-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="4b62c-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="4b62c-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="4b62c-113">id attribute</span></span>

<span data-ttu-id="4b62c-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="4b62c-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4b62c-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4b62c-118">Child elements</span></span>

|  <span data-ttu-id="4b62c-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b62c-119">Element</span></span> |  <span data-ttu-id="4b62c-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4b62c-120">Required</span></span>  |  <span data-ttu-id="4b62c-121">Описание</span><span class="sxs-lookup"><span data-stu-id="4b62c-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4b62c-122">Label</span><span class="sxs-lookup"><span data-stu-id="4b62c-122">Label</span></span>](#label)      | <span data-ttu-id="4b62c-123">Да</span><span class="sxs-lookup"><span data-stu-id="4b62c-123">Yes</span></span> |  <span data-ttu-id="4b62c-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="4b62c-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="4b62c-125">Icon</span><span class="sxs-lookup"><span data-stu-id="4b62c-125">Icon</span></span>](icon.md)      | <span data-ttu-id="4b62c-126">Да</span><span class="sxs-lookup"><span data-stu-id="4b62c-126">Yes</span></span> |  <span data-ttu-id="4b62c-127">Изображение для группы.</span><span class="sxs-lookup"><span data-stu-id="4b62c-127">The image for a group.</span></span> <span data-ttu-id="4b62c-128">Не поддерживается Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-128">Not supported in Outlook add-ins.</span></span> |
|  [<span data-ttu-id="4b62c-129">Control</span><span class="sxs-lookup"><span data-stu-id="4b62c-129">Control</span></span>](#control)    | <span data-ttu-id="4b62c-130">Нет</span><span class="sxs-lookup"><span data-stu-id="4b62c-130">No</span></span> |  <span data-ttu-id="4b62c-131">Представляет объект Control.</span><span class="sxs-lookup"><span data-stu-id="4b62c-131">Represents a Control object.</span></span> <span data-ttu-id="4b62c-132">Может быть ноль или больше.</span><span class="sxs-lookup"><span data-stu-id="4b62c-132">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="4b62c-133">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="4b62c-133">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="4b62c-134">Нет</span><span class="sxs-lookup"><span data-stu-id="4b62c-134">No</span></span> | <span data-ttu-id="4b62c-135">Представляет один из встроенных элементов Office элементов управления.</span><span class="sxs-lookup"><span data-stu-id="4b62c-135">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="4b62c-136">Может быть ноль или больше.</span><span class="sxs-lookup"><span data-stu-id="4b62c-136">Can be zero or more.</span></span> <span data-ttu-id="4b62c-137">Не поддерживается Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-137">Not supported in Outlook add-ins.</span></span>|
|  [<span data-ttu-id="4b62c-138">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="4b62c-138">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="4b62c-139">Нет</span><span class="sxs-lookup"><span data-stu-id="4b62c-139">No</span></span> |  <span data-ttu-id="4b62c-140">Указывает, должна ли группа отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-140">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="4b62c-141">Не поддерживается Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-141">Not supported in Outlook add-ins.</span></span> |

### <a name="label"></a><span data-ttu-id="4b62c-142">Label</span><span class="sxs-lookup"><span data-stu-id="4b62c-142">Label</span></span>

<span data-ttu-id="4b62c-143">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="4b62c-143">Required.</span></span> <span data-ttu-id="4b62c-144">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="4b62c-144">The label of the group.</span></span> <span data-ttu-id="4b62c-145">Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="4b62c-145">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="4b62c-146">Icon</span><span class="sxs-lookup"><span data-stu-id="4b62c-146">Icon</span></span>

<span data-ttu-id="4b62c-147">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="4b62c-147">Required.</span></span> <span data-ttu-id="4b62c-148">Если вкладка содержит большое количество групп и окно программы повторно, указанное изображение может отображаться вместо этого.</span><span class="sxs-lookup"><span data-stu-id="4b62c-148">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

> [!NOTE]
> <span data-ttu-id="4b62c-149">Этот элемент не поддерживается Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-149">This child element is not supported in Outlook add-ins.</span></span>

### <a name="control"></a><span data-ttu-id="4b62c-150">Средство контроля</span><span class="sxs-lookup"><span data-stu-id="4b62c-150">Control</span></span>

<span data-ttu-id="4b62c-151">Необязательный, но если его нет, то должен быть хотя бы один **OfficeControl.**</span><span class="sxs-lookup"><span data-stu-id="4b62c-151">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="4b62c-152">Сведения о типах поддерживаемых элементов управления см. в [элементе Control.](control.md)</span><span class="sxs-lookup"><span data-stu-id="4b62c-152">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="4b62c-153">Порядок  управления и **OfficeControl** в манифесте взаимозаменяем, и они могут быть взаимозаменяемыми, если существует несколько элементов, но все они должны быть ниже элемента **Icon.**</span><span class="sxs-lookup"><span data-stu-id="4b62c-153">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="officecontrol"></a><span data-ttu-id="4b62c-154">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="4b62c-154">OfficeControl</span></span>

<span data-ttu-id="4b62c-155">Необязательный, но если нет, должен быть хотя бы один **контроль.**</span><span class="sxs-lookup"><span data-stu-id="4b62c-155">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="4b62c-156">Включай один или несколько встроенных элементов Office в группу с `<OfficeControl>` элементами.</span><span class="sxs-lookup"><span data-stu-id="4b62c-156">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="4b62c-157">Атрибут указывает ID встроенного Office `id` управления.</span><span class="sxs-lookup"><span data-stu-id="4b62c-157">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="4b62c-158">Чтобы найти ID элементов управления, см. в рублях [Find the IDs of controls and control groups.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="4b62c-158">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="4b62c-159">Порядок  управления и **OfficeControl** в манифесте взаимозаменяем, и они могут быть взаимозаменяемыми, если существует несколько элементов, но все они должны быть ниже элемента **Icon.**</span><span class="sxs-lookup"><span data-stu-id="4b62c-159">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

> [!NOTE]
> <span data-ttu-id="4b62c-160">Этот элемент не поддерживается Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-160">This child element is not supported in Outlook add-ins.</span></span>

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

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="4b62c-161">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="4b62c-161">OverriddenByRibbonApi</span></span>

<span data-ttu-id="4b62c-162">Необязательный (boolean).</span><span class="sxs-lookup"><span data-stu-id="4b62c-162">Optional (boolean).</span></span> <span data-ttu-id="4b62c-163">Указывает, будет ли **группа** скрыта в сочетаниях приложений и платформ, поддерживаюх API, который устанавливает настраиваемую контекстную вкладку на ленту во время запуска.</span><span class="sxs-lookup"><span data-stu-id="4b62c-163">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="4b62c-164">Значение по умолчанию, если не присутствует, `false` является .</span><span class="sxs-lookup"><span data-stu-id="4b62c-164">The default value, if not present, is `false`.</span></span> <span data-ttu-id="4b62c-165">Если используется, **OverriddenByRibbonApi** должен быть *первым* ребенком **группы**.</span><span class="sxs-lookup"><span data-stu-id="4b62c-165">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="4b62c-166">Дополнительные сведения см. в [веб-сведениях OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="4b62c-166">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!NOTE]
> <span data-ttu-id="4b62c-167">Этот элемент не поддерживается Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b62c-167">This child element is not supported in Outlook add-ins.</span></span>

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
