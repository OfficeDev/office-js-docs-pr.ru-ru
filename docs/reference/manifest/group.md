---
title: Элемент Group в файле манифеста
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: ad1a566e259188ed20032bc5a3004736474e1f01
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670134"
---
# <a name="group-element"></a><span data-ttu-id="ea024-102">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="ea024-102">Group element</span></span>

<span data-ttu-id="ea024-p101">Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="ea024-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="ea024-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ea024-106">Attributes</span></span>

|  <span data-ttu-id="ea024-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="ea024-107">Attribute</span></span>  |  <span data-ttu-id="ea024-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ea024-108">Required</span></span>  |  <span data-ttu-id="ea024-109">Описание</span><span class="sxs-lookup"><span data-stu-id="ea024-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ea024-110">id</span><span class="sxs-lookup"><span data-stu-id="ea024-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="ea024-111">Да</span><span class="sxs-lookup"><span data-stu-id="ea024-111">Yes</span></span>  | <span data-ttu-id="ea024-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="ea024-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="ea024-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="ea024-113">id attribute</span></span>

<span data-ttu-id="ea024-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="ea024-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ea024-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ea024-118">Child elements</span></span>
|  <span data-ttu-id="ea024-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="ea024-119">Element</span></span> |  <span data-ttu-id="ea024-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ea024-120">Required</span></span>  |  <span data-ttu-id="ea024-121">Описание</span><span class="sxs-lookup"><span data-stu-id="ea024-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ea024-122">Label</span><span class="sxs-lookup"><span data-stu-id="ea024-122">Label</span></span>](#label)      | <span data-ttu-id="ea024-123">Да</span><span class="sxs-lookup"><span data-stu-id="ea024-123">Yes</span></span> |  <span data-ttu-id="ea024-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="ea024-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="ea024-125">Control</span><span class="sxs-lookup"><span data-stu-id="ea024-125">Control</span></span>](#control)    | <span data-ttu-id="ea024-126">Да</span><span class="sxs-lookup"><span data-stu-id="ea024-126">Yes</span></span> |  <span data-ttu-id="ea024-127">Коллекция одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="ea024-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="ea024-128">Label</span><span class="sxs-lookup"><span data-stu-id="ea024-128">Label</span></span> 

<span data-ttu-id="ea024-p103">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ea024-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="ea024-132">Control</span><span class="sxs-lookup"><span data-stu-id="ea024-132">Control</span></span>
<span data-ttu-id="ea024-133">В группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="ea024-133">A group requires at least one control.</span></span> <span data-ttu-id="ea024-134">Дополнительные сведения о поддерживаемых типах элементов управления приведены в элементе [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="ea024-134">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
