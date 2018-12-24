---
title: Элемент Group в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 13cd9bbe6f602fd1779caea487e34177c3e9d483
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433703"
---
# <a name="group-element"></a><span data-ttu-id="c7bd1-102">Элемент Group</span><span class="sxs-lookup"><span data-stu-id="c7bd1-102">Group element</span></span>

<span data-ttu-id="c7bd1-p101">Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="c7bd1-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="c7bd1-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c7bd1-106">Attributes</span></span>

|  <span data-ttu-id="c7bd1-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c7bd1-107">Attribute</span></span>  |  <span data-ttu-id="c7bd1-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c7bd1-108">Required</span></span>  |  <span data-ttu-id="c7bd1-109">Описание</span><span class="sxs-lookup"><span data-stu-id="c7bd1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c7bd1-110">id</span><span class="sxs-lookup"><span data-stu-id="c7bd1-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="c7bd1-111">Да</span><span class="sxs-lookup"><span data-stu-id="c7bd1-111">Yes</span></span>  | <span data-ttu-id="c7bd1-112">Уникальный идентификатор группы.</span><span class="sxs-lookup"><span data-stu-id="c7bd1-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="c7bd1-113">Атрибут id</span><span class="sxs-lookup"><span data-stu-id="c7bd1-113">id attribute</span></span>

<span data-ttu-id="c7bd1-p102">Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.</span><span class="sxs-lookup"><span data-stu-id="c7bd1-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c7bd1-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c7bd1-118">Child elements</span></span>
|  <span data-ttu-id="c7bd1-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="c7bd1-119">Element</span></span> |  <span data-ttu-id="c7bd1-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c7bd1-120">Required</span></span>  |  <span data-ttu-id="c7bd1-121">Описание</span><span class="sxs-lookup"><span data-stu-id="c7bd1-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c7bd1-122">Label</span><span class="sxs-lookup"><span data-stu-id="c7bd1-122">Label</span></span>](#label)      | <span data-ttu-id="c7bd1-123">Да</span><span class="sxs-lookup"><span data-stu-id="c7bd1-123">Yes</span></span> |  <span data-ttu-id="c7bd1-124">Метка элемента CustomTab или группы.</span><span class="sxs-lookup"><span data-stu-id="c7bd1-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="c7bd1-125">Control</span><span class="sxs-lookup"><span data-stu-id="c7bd1-125">Control</span></span>](#control)    | <span data-ttu-id="c7bd1-126">Да</span><span class="sxs-lookup"><span data-stu-id="c7bd1-126">Yes</span></span> |  <span data-ttu-id="c7bd1-127">Коллекция одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="c7bd1-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="c7bd1-128">Label</span><span class="sxs-lookup"><span data-stu-id="c7bd1-128">Label</span></span> 

<span data-ttu-id="c7bd1-p103">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c7bd1-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="c7bd1-132">Control</span><span class="sxs-lookup"><span data-stu-id="c7bd1-132">Control</span></span>
<span data-ttu-id="c7bd1-133">Для группы требуется по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="c7bd1-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```