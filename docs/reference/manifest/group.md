---
title: Элемент Group в файле манифеста
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 35db4829b40078e97fbfc007e2fb552e00875f9c
ms.sourcegitcommit: 164b11b1e9d2ae20b3d816092025b32a9070450f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2019
ms.locfileid: "39818729"
---
# <a name="group-element"></a><span data-ttu-id="e2c27-102">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="e2c27-102">Group element</span></span>

<span data-ttu-id="e2c27-p101">Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="e2c27-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="e2c27-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e2c27-106">Attributes</span></span>

|  <span data-ttu-id="e2c27-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e2c27-107">Attribute</span></span>  |  <span data-ttu-id="e2c27-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e2c27-108">Required</span></span>  |  <span data-ttu-id="e2c27-109">Описание</span><span class="sxs-lookup"><span data-stu-id="e2c27-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e2c27-110">id</span><span class="sxs-lookup"><span data-stu-id="e2c27-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="e2c27-111">Да</span><span class="sxs-lookup"><span data-stu-id="e2c27-111">Yes</span></span>  | <span data-ttu-id="e2c27-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="e2c27-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="e2c27-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="e2c27-113">id attribute</span></span>

<span data-ttu-id="e2c27-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="e2c27-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e2c27-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e2c27-118">Child elements</span></span>
|  <span data-ttu-id="e2c27-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="e2c27-119">Element</span></span> |  <span data-ttu-id="e2c27-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e2c27-120">Required</span></span>  |  <span data-ttu-id="e2c27-121">Описание</span><span class="sxs-lookup"><span data-stu-id="e2c27-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e2c27-122">Label</span><span class="sxs-lookup"><span data-stu-id="e2c27-122">Label</span></span>](#label)      | <span data-ttu-id="e2c27-123">Да</span><span class="sxs-lookup"><span data-stu-id="e2c27-123">Yes</span></span> |  <span data-ttu-id="e2c27-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="e2c27-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="e2c27-125">Icon</span><span class="sxs-lookup"><span data-stu-id="e2c27-125">Icon</span></span>](icon.md)      | <span data-ttu-id="e2c27-126">Да</span><span class="sxs-lookup"><span data-stu-id="e2c27-126">Yes</span></span> |  <span data-ttu-id="e2c27-127">Изображение для группы.</span><span class="sxs-lookup"><span data-stu-id="e2c27-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="e2c27-128">Control</span><span class="sxs-lookup"><span data-stu-id="e2c27-128">Control</span></span>](#control)    | <span data-ttu-id="e2c27-129">Да</span><span class="sxs-lookup"><span data-stu-id="e2c27-129">Yes</span></span> |  <span data-ttu-id="e2c27-130">Коллекция одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="e2c27-130">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="e2c27-131">Label</span><span class="sxs-lookup"><span data-stu-id="e2c27-131">Label</span></span> 

<span data-ttu-id="e2c27-p103">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e2c27-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="e2c27-135">Icon</span><span class="sxs-lookup"><span data-stu-id="e2c27-135">Icon</span></span>

<span data-ttu-id="e2c27-136">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="e2c27-136">Required.</span></span> <span data-ttu-id="e2c27-137">Если вкладка содержит большое количество групп и изменяется размер окна программы, вместо этого может отображаться указанное изображение.</span><span class="sxs-lookup"><span data-stu-id="e2c27-137">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="e2c27-138">Control</span><span class="sxs-lookup"><span data-stu-id="e2c27-138">Control</span></span>
<span data-ttu-id="e2c27-139">В группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="e2c27-139">A group requires at least one control.</span></span> <span data-ttu-id="e2c27-140">Дополнительные сведения о поддерживаемых типах элементов управления приведены в элементе [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="e2c27-140">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
