---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: a598232f230a120dccd58024e760c2172a769727
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611829"
---
# <a name="group-element"></a><span data-ttu-id="4d09b-103">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="4d09b-103">Group element</span></span>

<span data-ttu-id="4d09b-p101">Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="4d09b-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="4d09b-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4d09b-107">Attributes</span></span>

|  <span data-ttu-id="4d09b-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="4d09b-108">Attribute</span></span>  |  <span data-ttu-id="4d09b-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4d09b-109">Required</span></span>  |  <span data-ttu-id="4d09b-110">Описание</span><span class="sxs-lookup"><span data-stu-id="4d09b-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4d09b-111">id</span><span class="sxs-lookup"><span data-stu-id="4d09b-111">id</span></span>](#id-attribute)  |  <span data-ttu-id="4d09b-112">Да</span><span class="sxs-lookup"><span data-stu-id="4d09b-112">Yes</span></span>  | <span data-ttu-id="4d09b-113">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="4d09b-113">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="4d09b-114">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="4d09b-114">id attribute</span></span>

<span data-ttu-id="4d09b-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="4d09b-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4d09b-119">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4d09b-119">Child elements</span></span>
|  <span data-ttu-id="4d09b-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="4d09b-120">Element</span></span> |  <span data-ttu-id="4d09b-121">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4d09b-121">Required</span></span>  |  <span data-ttu-id="4d09b-122">Описание</span><span class="sxs-lookup"><span data-stu-id="4d09b-122">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4d09b-123">Label</span><span class="sxs-lookup"><span data-stu-id="4d09b-123">Label</span></span>](#label)      | <span data-ttu-id="4d09b-124">Да</span><span class="sxs-lookup"><span data-stu-id="4d09b-124">Yes</span></span> |  <span data-ttu-id="4d09b-125">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="4d09b-125">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="4d09b-126">Icon</span><span class="sxs-lookup"><span data-stu-id="4d09b-126">Icon</span></span>](icon.md)      | <span data-ttu-id="4d09b-127">Да</span><span class="sxs-lookup"><span data-stu-id="4d09b-127">Yes</span></span> |  <span data-ttu-id="4d09b-128">Изображение для группы.</span><span class="sxs-lookup"><span data-stu-id="4d09b-128">The image for a group.</span></span>  |
|  [<span data-ttu-id="4d09b-129">Control</span><span class="sxs-lookup"><span data-stu-id="4d09b-129">Control</span></span>](#control)    | <span data-ttu-id="4d09b-130">Да</span><span class="sxs-lookup"><span data-stu-id="4d09b-130">Yes</span></span> |  <span data-ttu-id="4d09b-131">Коллекция одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="4d09b-131">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="4d09b-132">Label</span><span class="sxs-lookup"><span data-stu-id="4d09b-132">Label</span></span> 

<span data-ttu-id="4d09b-133">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="4d09b-133">Required.</span></span> <span data-ttu-id="4d09b-134">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="4d09b-134">The label of the group.</span></span> <span data-ttu-id="4d09b-135">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="4d09b-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="4d09b-136">Icon</span><span class="sxs-lookup"><span data-stu-id="4d09b-136">Icon</span></span>

<span data-ttu-id="4d09b-137">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="4d09b-137">Required.</span></span> <span data-ttu-id="4d09b-138">Если вкладка содержит большое количество групп и изменяется размер окна программы, вместо этого может отображаться указанное изображение.</span><span class="sxs-lookup"><span data-stu-id="4d09b-138">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="4d09b-139">Control</span><span class="sxs-lookup"><span data-stu-id="4d09b-139">Control</span></span>
<span data-ttu-id="4d09b-140">В группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="4d09b-140">A group requires at least one control.</span></span> <span data-ttu-id="4d09b-141">Дополнительные сведения о поддерживаемых типах элементов управления приведены в элементе [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="4d09b-141">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
