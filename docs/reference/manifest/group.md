---
title: Элемент Group в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7cc1f4c398eeb013eb6033b207b395466f7d72ca
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450711"
---
# <a name="group-element"></a><span data-ttu-id="4eafc-102">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="4eafc-102">Group element</span></span>

<span data-ttu-id="4eafc-p101">Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="4eafc-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="4eafc-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4eafc-106">Attributes</span></span>

|  <span data-ttu-id="4eafc-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="4eafc-107">Attribute</span></span>  |  <span data-ttu-id="4eafc-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4eafc-108">Required</span></span>  |  <span data-ttu-id="4eafc-109">Описание</span><span class="sxs-lookup"><span data-stu-id="4eafc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4eafc-110">id</span><span class="sxs-lookup"><span data-stu-id="4eafc-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="4eafc-111">Да</span><span class="sxs-lookup"><span data-stu-id="4eafc-111">Yes</span></span>  | <span data-ttu-id="4eafc-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="4eafc-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="4eafc-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="4eafc-113">id attribute</span></span>

<span data-ttu-id="4eafc-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="4eafc-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4eafc-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4eafc-118">Child elements</span></span>
|  <span data-ttu-id="4eafc-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="4eafc-119">Element</span></span> |  <span data-ttu-id="4eafc-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4eafc-120">Required</span></span>  |  <span data-ttu-id="4eafc-121">Описание</span><span class="sxs-lookup"><span data-stu-id="4eafc-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4eafc-122">Label</span><span class="sxs-lookup"><span data-stu-id="4eafc-122">Label</span></span>](#label)      | <span data-ttu-id="4eafc-123">Да</span><span class="sxs-lookup"><span data-stu-id="4eafc-123">Yes</span></span> |  <span data-ttu-id="4eafc-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="4eafc-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="4eafc-125">Control</span><span class="sxs-lookup"><span data-stu-id="4eafc-125">Control</span></span>](#control)    | <span data-ttu-id="4eafc-126">Да</span><span class="sxs-lookup"><span data-stu-id="4eafc-126">Yes</span></span> |  <span data-ttu-id="4eafc-127">Коллекция одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="4eafc-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="4eafc-128">Label</span><span class="sxs-lookup"><span data-stu-id="4eafc-128">Label</span></span> 

<span data-ttu-id="4eafc-p103">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="4eafc-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="4eafc-132">Control</span><span class="sxs-lookup"><span data-stu-id="4eafc-132">Control</span></span>
<span data-ttu-id="4eafc-133">В группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="4eafc-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
