---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 99670b27d963060a008899a8808ca967cfd710a6
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087940"
---
# <a name="customtab-element"></a><span data-ttu-id="33d37-103">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="33d37-103">CustomTab element</span></span>

<span data-ttu-id="33d37-104">На ленте укажите вкладку и группу для команд надстройки.</span><span class="sxs-lookup"><span data-stu-id="33d37-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="33d37-105">Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="33d37-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="33d37-106">На пользовательских вкладках надстройка может иметь настраиваемые или встроенные группы.</span><span class="sxs-lookup"><span data-stu-id="33d37-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="33d37-107">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="33d37-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="33d37-108">Атрибут **ID** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="33d37-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33d37-109">В Outlook на Mac `CustomTab` элемент недоступен, поэтому необходимо использовать [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="33d37-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="33d37-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="33d37-110">Child elements</span></span>

|  <span data-ttu-id="33d37-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="33d37-111">Element</span></span> |  <span data-ttu-id="33d37-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="33d37-112">Required</span></span>  |  <span data-ttu-id="33d37-113">Описание</span><span class="sxs-lookup"><span data-stu-id="33d37-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="33d37-114">Group</span><span class="sxs-lookup"><span data-stu-id="33d37-114">Group</span></span>](group.md)      | <span data-ttu-id="33d37-115">Нет</span><span class="sxs-lookup"><span data-stu-id="33d37-115">No</span></span> |  <span data-ttu-id="33d37-116">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="33d37-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="33d37-117">оффицеграуп</span><span class="sxs-lookup"><span data-stu-id="33d37-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="33d37-118">Нет</span><span class="sxs-lookup"><span data-stu-id="33d37-118">No</span></span> |  <span data-ttu-id="33d37-119">Представляет встроенную группу управления Office.</span><span class="sxs-lookup"><span data-stu-id="33d37-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="33d37-120">Label</span><span class="sxs-lookup"><span data-stu-id="33d37-120">Label</span></span>](#label-tab)      | <span data-ttu-id="33d37-121">Да</span><span class="sxs-lookup"><span data-stu-id="33d37-121">Yes</span></span> |  <span data-ttu-id="33d37-122">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="33d37-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="33d37-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="33d37-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="33d37-124">Нет</span><span class="sxs-lookup"><span data-stu-id="33d37-124">No</span></span> |  <span data-ttu-id="33d37-125">Указывает, что настраиваемая вкладка должна находиться сразу после указанной встроенной вкладки Office.</span><span class="sxs-lookup"><span data-stu-id="33d37-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="33d37-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="33d37-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="33d37-127">Нет</span><span class="sxs-lookup"><span data-stu-id="33d37-127">No</span></span> |  <span data-ttu-id="33d37-128">Указывает, что настраиваемая вкладка должна находиться непосредственно перед указанной встроенной вкладкой Office.</span><span class="sxs-lookup"><span data-stu-id="33d37-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="33d37-129">Group</span><span class="sxs-lookup"><span data-stu-id="33d37-129">Group</span></span>

<span data-ttu-id="33d37-130">Необязательный параметр, но если он отсутствует, должен быть по крайней мере один элемент **оффицеграуп** .</span><span class="sxs-lookup"><span data-stu-id="33d37-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="33d37-131">Просмотрите [элемент Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="33d37-131">See [Group element](group.md).</span></span> <span data-ttu-id="33d37-132">Порядок **группировки** и **оффицеграуп** в манифесте должен быть указан в том порядке, в котором они должны отображаться на вкладке Настраиваемый. Они могут быть интерминглед, если имеется несколько элементов, но они должны быть над элементом **Label** .</span><span class="sxs-lookup"><span data-stu-id="33d37-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="33d37-133">оффицеграуп</span><span class="sxs-lookup"><span data-stu-id="33d37-133">OfficeGroup</span></span>

<span data-ttu-id="33d37-134">Необязательный параметр, но если он отсутствует, должен быть хотя бы один элемент **Group** .</span><span class="sxs-lookup"><span data-stu-id="33d37-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="33d37-135">Представляет встроенную группу управления Office.</span><span class="sxs-lookup"><span data-stu-id="33d37-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="33d37-136">Атрибут **ID** указывает идентификатор встроенной группы Office.</span><span class="sxs-lookup"><span data-stu-id="33d37-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="33d37-137">Чтобы найти идентификатор встроенной группы, ознакомьтесь со статьей [Поиск идентификаторов элементов управления и групп элементов управления](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="33d37-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="33d37-138">Порядок **группировки** и **оффицеграуп** в манифесте должен быть указан в том порядке, в котором они должны отображаться на вкладке Настраиваемый. Они могут быть интерминглед, если имеется несколько элементов, но они должны быть над элементом **Label** .</span><span class="sxs-lookup"><span data-stu-id="33d37-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="33d37-139">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="33d37-139">Label (Tab)</span></span>

<span data-ttu-id="33d37-140">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="33d37-140">Required.</span></span> <span data-ttu-id="33d37-141">Метка настраиваемой вкладки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="33d37-141">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="33d37-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="33d37-142">InsertAfter</span></span>

<span data-ttu-id="33d37-143">Необязательный атрибут.</span><span class="sxs-lookup"><span data-stu-id="33d37-143">Optional.</span></span> <span data-ttu-id="33d37-144">Указывает, что настраиваемая вкладка должна находиться сразу после указанной встроенной вкладки Office. Значение элемента — это идентификатор встроенной вкладки, например "Табхоме" или "Табревиев".</span><span class="sxs-lookup"><span data-stu-id="33d37-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="33d37-145">(См. раздел [Поиск идентификаторов элементов управления и групп элементов управления](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) Если этот параметр указан, он должен находиться после элемента **Label** .</span><span class="sxs-lookup"><span data-stu-id="33d37-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="33d37-146">Невозможно использовать одновременно **InsertAfter** и **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="33d37-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="33d37-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="33d37-147">InsertBefore</span></span>

<span data-ttu-id="33d37-148">Необязательный атрибут.</span><span class="sxs-lookup"><span data-stu-id="33d37-148">Optional.</span></span> <span data-ttu-id="33d37-149">Указывает, что настраиваемая вкладка должна находиться непосредственно перед указанной встроенной вкладкой Office. Значение элемента — это идентификатор встроенной вкладки, например "Табхоме" или "Табревиев".</span><span class="sxs-lookup"><span data-stu-id="33d37-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="33d37-150">(См. раздел [Поиск идентификаторов элементов управления и групп элементов управления](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  Если этот параметр указан, он должен находиться после элемента **Label** .</span><span class="sxs-lookup"><span data-stu-id="33d37-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="33d37-151">Невозможно использовать одновременно **InsertAfter** и **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="33d37-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="33d37-152">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="33d37-152">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
