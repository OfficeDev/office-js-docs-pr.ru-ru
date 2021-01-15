---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771328"
---
# <a name="customtab-element"></a><span data-ttu-id="03cab-103">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="03cab-103">CustomTab element</span></span>

<span data-ttu-id="03cab-104">На ленте укажите вкладку и группу для команд надстройки.</span><span class="sxs-lookup"><span data-stu-id="03cab-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="03cab-105">Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="03cab-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="03cab-106">На настраиваемой вкладке надстройка может иметь настраиваемые или встроенные группы.</span><span class="sxs-lookup"><span data-stu-id="03cab-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="03cab-107">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="03cab-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="03cab-108">Атрибут **id** должен быть уникальным в манифесте.</span><span class="sxs-lookup"><span data-stu-id="03cab-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="03cab-109">В Outlook для Mac элемент не доступен, поэтому придется `CustomTab` использовать [OfficeTab.](officetab.md)</span><span class="sxs-lookup"><span data-stu-id="03cab-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="03cab-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="03cab-110">Child elements</span></span>

|  <span data-ttu-id="03cab-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="03cab-111">Element</span></span> |  <span data-ttu-id="03cab-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="03cab-112">Required</span></span>  |  <span data-ttu-id="03cab-113">Описание</span><span class="sxs-lookup"><span data-stu-id="03cab-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="03cab-114">Group</span><span class="sxs-lookup"><span data-stu-id="03cab-114">Group</span></span>](group.md)      | <span data-ttu-id="03cab-115">Нет</span><span class="sxs-lookup"><span data-stu-id="03cab-115">No</span></span> |  <span data-ttu-id="03cab-116">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="03cab-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="03cab-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="03cab-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="03cab-118">Нет</span><span class="sxs-lookup"><span data-stu-id="03cab-118">No</span></span> |  <span data-ttu-id="03cab-119">Представляет встроенную группу управления Office.</span><span class="sxs-lookup"><span data-stu-id="03cab-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="03cab-120">Label</span><span class="sxs-lookup"><span data-stu-id="03cab-120">Label</span></span>](#label-tab)      | <span data-ttu-id="03cab-121">Да</span><span class="sxs-lookup"><span data-stu-id="03cab-121">Yes</span></span> |  <span data-ttu-id="03cab-122">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="03cab-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="03cab-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="03cab-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="03cab-124">Нет</span><span class="sxs-lookup"><span data-stu-id="03cab-124">No</span></span> |  <span data-ttu-id="03cab-125">Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office.</span><span class="sxs-lookup"><span data-stu-id="03cab-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="03cab-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="03cab-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="03cab-127">Нет</span><span class="sxs-lookup"><span data-stu-id="03cab-127">No</span></span> |  <span data-ttu-id="03cab-128">Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office.</span><span class="sxs-lookup"><span data-stu-id="03cab-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="03cab-129">Группа</span><span class="sxs-lookup"><span data-stu-id="03cab-129">Group</span></span>

<span data-ttu-id="03cab-130">Необязательный, но если его нет, должен быть хотя бы один **элемент OfficeGroup.**</span><span class="sxs-lookup"><span data-stu-id="03cab-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="03cab-131">См. [элемент Group.](group.md)</span><span class="sxs-lookup"><span data-stu-id="03cab-131">See [Group element](group.md).</span></span> <span data-ttu-id="03cab-132">Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Они могут быть перемелены, если существует несколько элементов, но все они должны быть над **элементом Label.**</span><span class="sxs-lookup"><span data-stu-id="03cab-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="03cab-133">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="03cab-133">OfficeGroup</span></span>

<span data-ttu-id="03cab-134">Необязательный, но если его нет, должен быть хотя бы один **элемент Group.**</span><span class="sxs-lookup"><span data-stu-id="03cab-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="03cab-135">Представляет встроенную группу управления Office.</span><span class="sxs-lookup"><span data-stu-id="03cab-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="03cab-136">Атрибут **id** указывает ИД встроенной группы Office.</span><span class="sxs-lookup"><span data-stu-id="03cab-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="03cab-137">Чтобы найти ИД встроенной группы, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="03cab-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="03cab-138">Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Они могут быть перемелены, если существует несколько элементов, но все они должны быть над **элементом Label.**</span><span class="sxs-lookup"><span data-stu-id="03cab-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="03cab-139">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="03cab-139">Label (Tab)</span></span>

<span data-ttu-id="03cab-140">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="03cab-140">Required.</span></span> <span data-ttu-id="03cab-141">Метка пользовательской вкладки. Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="03cab-141">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="03cab-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="03cab-142">InsertAfter</span></span>

<span data-ttu-id="03cab-143">Необязательное свойство.</span><span class="sxs-lookup"><span data-stu-id="03cab-143">Optional.</span></span> <span data-ttu-id="03cab-144">Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview.</span><span class="sxs-lookup"><span data-stu-id="03cab-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="03cab-145">[(См. "Поиск ИД элементов управления и групп элементов управления".](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Если этот элемент заметок, он должен быть после **элемента Label.**</span><span class="sxs-lookup"><span data-stu-id="03cab-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="03cab-146">Невозможно одновременно **insertAfter** и **InsertBefore.**</span><span class="sxs-lookup"><span data-stu-id="03cab-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="03cab-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="03cab-147">InsertBefore</span></span>

<span data-ttu-id="03cab-148">Необязательное свойство.</span><span class="sxs-lookup"><span data-stu-id="03cab-148">Optional.</span></span> <span data-ttu-id="03cab-149">Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview.</span><span class="sxs-lookup"><span data-stu-id="03cab-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="03cab-150">[(См. "Поиск ИД элементов управления и групп элементов управления".](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Если этот элемент заметок, он должен быть после **элемента Label.**</span><span class="sxs-lookup"><span data-stu-id="03cab-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="03cab-151">Невозможно одновременно **insertAfter** и **InsertBefore.**</span><span class="sxs-lookup"><span data-stu-id="03cab-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="03cab-152">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="03cab-152">CustomTab example</span></span>

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
