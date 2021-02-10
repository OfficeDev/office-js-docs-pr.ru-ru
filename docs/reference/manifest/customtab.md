---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173929"
---
# <a name="customtab-element"></a><span data-ttu-id="9d24b-103">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="9d24b-103">CustomTab element</span></span>

<span data-ttu-id="9d24b-104">На ленте укажите вкладку и группу для команд надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d24b-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="9d24b-105">Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="9d24b-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="9d24b-106">На настраиваемой вкладке надстройка может иметь настраиваемые или встроенные группы.</span><span class="sxs-lookup"><span data-stu-id="9d24b-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="9d24b-107">Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="9d24b-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="9d24b-108">Атрибут **id** должен быть уникальным в манифесте.</span><span class="sxs-lookup"><span data-stu-id="9d24b-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d24b-109">В Outlook для Mac элемент не доступен, поэтому необходимо использовать `CustomTab` [OfficeTab.](officetab.md)</span><span class="sxs-lookup"><span data-stu-id="9d24b-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9d24b-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9d24b-110">Child elements</span></span>

|  <span data-ttu-id="9d24b-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="9d24b-111">Element</span></span> |  <span data-ttu-id="9d24b-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9d24b-112">Required</span></span>  |  <span data-ttu-id="9d24b-113">Описание</span><span class="sxs-lookup"><span data-stu-id="9d24b-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9d24b-114">Group</span><span class="sxs-lookup"><span data-stu-id="9d24b-114">Group</span></span>](group.md)      | <span data-ttu-id="9d24b-115">Нет</span><span class="sxs-lookup"><span data-stu-id="9d24b-115">No</span></span> |  <span data-ttu-id="9d24b-116">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="9d24b-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="9d24b-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="9d24b-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="9d24b-118">Нет</span><span class="sxs-lookup"><span data-stu-id="9d24b-118">No</span></span> |  <span data-ttu-id="9d24b-119">Представляет встроенную группу управления Office.</span><span class="sxs-lookup"><span data-stu-id="9d24b-119">Represents a built-in Office control group.</span></span> <span data-ttu-id="9d24b-120">**Важно!** Отсутствует в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-120">**Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="9d24b-121">Label</span><span class="sxs-lookup"><span data-stu-id="9d24b-121">Label</span></span>](#label-tab)      | <span data-ttu-id="9d24b-122">Да</span><span class="sxs-lookup"><span data-stu-id="9d24b-122">Yes</span></span> |  <span data-ttu-id="9d24b-123">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="9d24b-123">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="9d24b-124">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="9d24b-124">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="9d24b-125">Нет</span><span class="sxs-lookup"><span data-stu-id="9d24b-125">No</span></span> |  <span data-ttu-id="9d24b-126">Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки **Office.** Важно: отсутствует в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-126">Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="9d24b-127">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="9d24b-127">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="9d24b-128">Нет</span><span class="sxs-lookup"><span data-stu-id="9d24b-128">No</span></span> |  <span data-ttu-id="9d24b-129">Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке **Office.** Важно! Отсутствует в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-129">Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="9d24b-130">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="9d24b-130">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="9d24b-131">Нет</span><span class="sxs-lookup"><span data-stu-id="9d24b-131">No</span></span> |  <span data-ttu-id="9d24b-132">Указывает, должна ли настраиваемая вкладка отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="9d24b-132">Specifies whether the custom tab should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="9d24b-133">**Важно!** Отсутствует в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-133">**Important**: Not available in Outlook.</span></span> |

### <a name="group"></a><span data-ttu-id="9d24b-134">Группа</span><span class="sxs-lookup"><span data-stu-id="9d24b-134">Group</span></span>

<span data-ttu-id="9d24b-135">Необязательный, но если его нет, должен быть хотя бы один **элемент OfficeGroup.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-135">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="9d24b-136">См. [элемент Group.](group.md)</span><span class="sxs-lookup"><span data-stu-id="9d24b-136">See [Group element](group.md).</span></span> <span data-ttu-id="9d24b-137">Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Их можно перемесить, если существует несколько элементов, но все они должны быть над **элементом Label.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-137">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="9d24b-138">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="9d24b-138">OfficeGroup</span></span>

<span data-ttu-id="9d24b-139">Необязательный, но если его нет, должен быть хотя бы один **элемент Group.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-139">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="9d24b-140">Представляет встроенную группу управления Office.</span><span class="sxs-lookup"><span data-stu-id="9d24b-140">Represents a built-in Office control group.</span></span> <span data-ttu-id="9d24b-141">Атрибут **id** указывает ИД встроенной группы Office.</span><span class="sxs-lookup"><span data-stu-id="9d24b-141">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="9d24b-142">Чтобы найти ИД встроенной группы, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="9d24b-142">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="9d24b-143">Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Их можно перемесить, если существует несколько элементов, но все они должны быть над **элементом Label.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-143">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d24b-144">Элемент `OfficeGroup` не доступен в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-144">The `OfficeGroup` element is not available in Outlook.</span></span>

### <a name="label-tab"></a><span data-ttu-id="9d24b-145">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="9d24b-145">Label (Tab)</span></span>

<span data-ttu-id="9d24b-146">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="9d24b-146">Required.</span></span> <span data-ttu-id="9d24b-147">Метка пользовательской вкладки. Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="9d24b-147">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="9d24b-148">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="9d24b-148">InsertAfter</span></span>

<span data-ttu-id="9d24b-149">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="9d24b-149">Optional.</span></span> <span data-ttu-id="9d24b-150">Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview.</span><span class="sxs-lookup"><span data-stu-id="9d24b-150">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="9d24b-151">[(См. поиск ИД элементов управления и групп элементов управления.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Если этот элемент заметим, он должен быть после **элемента Label.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-151">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="9d24b-152">Невозможно одновременное **добавление InsertAfter** и **InsertBefore.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-152">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d24b-153">Элемент `InsertAfter` не доступен в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-153">The `InsertAfter` element is not available in Outlook.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="9d24b-154">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="9d24b-154">InsertBefore</span></span>

<span data-ttu-id="9d24b-155">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="9d24b-155">Optional.</span></span> <span data-ttu-id="9d24b-156">Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview.</span><span class="sxs-lookup"><span data-stu-id="9d24b-156">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="9d24b-157">[(См. поиск ИД элементов управления и групп элементов управления.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Если этот элемент заметим, он должен быть после **элемента Label.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-157">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="9d24b-158">Невозможно одновременное **добавление InsertAfter** и **InsertBefore.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-158">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d24b-159">Элемент `InsertBefore` не доступен в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-159">The `InsertBefore` element is not available in Outlook.</span></span>

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="9d24b-160">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="9d24b-160">OverriddenByRibbonApi</span></span>

<span data-ttu-id="9d24b-161">Необязательный (boolean).</span><span class="sxs-lookup"><span data-stu-id="9d24b-161">Optional (boolean).</span></span> <span data-ttu-id="9d24b-162">Указывает, будет ли **customTab** скрыт в сочетаниях приложений и платформ, которые поддерживают API, устанавливая настраиваемую контекстную вкладку на ленту во время работы.</span><span class="sxs-lookup"><span data-stu-id="9d24b-162">Specifies whether the **CustomTab** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="9d24b-163">Значение по умолчанию (если его нет) `false` — .</span><span class="sxs-lookup"><span data-stu-id="9d24b-163">The default value, if not present, is `false`.</span></span> <span data-ttu-id="9d24b-164">Если используется, **OverriddenByRibbonApi**  должен быть первым child of **CustomTab.**</span><span class="sxs-lookup"><span data-stu-id="9d24b-164">If used, **OverriddenByRibbonApi** must be the *first* child of **CustomTab**.</span></span> <span data-ttu-id="9d24b-165">Дополнительные сведения [см. в подразделе OverriddenByRibbonApi.](overriddenbyribbonapi.md)</span><span class="sxs-lookup"><span data-stu-id="9d24b-165">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d24b-166">Элемент `OverriddenByRibbonApi` не доступен в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d24b-166">The `OverriddenByRibbonApi` element is not available in Outlook.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="9d24b-167">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="9d24b-167">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
