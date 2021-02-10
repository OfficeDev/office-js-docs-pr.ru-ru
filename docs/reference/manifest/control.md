---
title: Элемент Control в файле манифеста
description: Определяет функцию JavaScript, которая выполняет действие или открывает область задач.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 737902bef52edeb70e2c5760df5bb589b624271b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173985"
---
# <a name="control-element"></a><span data-ttu-id="eeeaa-103">Элемент Control</span><span class="sxs-lookup"><span data-stu-id="eeeaa-103">Control element</span></span>

<span data-ttu-id="eeeaa-p101">Определяет функцию JavaScript, которая выполняет действие или открывает область задач. Элемент **Control** может быть кнопкой или пунктом меню. Элемент [Group](group.md) должен содержать по крайней мере один элемент **Control**.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="eeeaa-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="eeeaa-107">Attributes</span></span>

|  <span data-ttu-id="eeeaa-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="eeeaa-108">Attribute</span></span>  |  <span data-ttu-id="eeeaa-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="eeeaa-109">Required</span></span>  |  <span data-ttu-id="eeeaa-110">Описание</span><span class="sxs-lookup"><span data-stu-id="eeeaa-110">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="eeeaa-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-111">**xsi:type**</span></span>|<span data-ttu-id="eeeaa-112">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-112">Yes</span></span>|<span data-ttu-id="eeeaa-p102">Тип определяемого элемента управления. Доступные варианты: `Button`, `Menu` или `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="eeeaa-115">**id**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-115">**id**</span></span>|<span data-ttu-id="eeeaa-116">Нет</span><span class="sxs-lookup"><span data-stu-id="eeeaa-116">No</span></span>|<span data-ttu-id="eeeaa-p103">ИД элемента управления. Может содержать до 125 знаков.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="eeeaa-119">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-119">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="eeeaa-120">Применяется только к элементам **Control**, которые содержатся в элементе [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="eeeaa-120">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="eeeaa-121">Элемент управления ''Кнопка''</span><span class="sxs-lookup"><span data-stu-id="eeeaa-121">Button control</span></span>

<span data-ttu-id="eeeaa-p105">Кнопка выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка" должен иметь элемент `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="eeeaa-125">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="eeeaa-125">Child elements</span></span>
|  <span data-ttu-id="eeeaa-126">Элемент</span><span class="sxs-lookup"><span data-stu-id="eeeaa-126">Element</span></span> |  <span data-ttu-id="eeeaa-127">Обязательный</span><span class="sxs-lookup"><span data-stu-id="eeeaa-127">Required</span></span>  |  <span data-ttu-id="eeeaa-128">Описание</span><span class="sxs-lookup"><span data-stu-id="eeeaa-128">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="eeeaa-129">**Label**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-129">**Label**</span></span>     | <span data-ttu-id="eeeaa-130">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-130">Yes</span></span> |  <span data-ttu-id="eeeaa-131">Текст для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-131">The text for the button.</span></span> <span data-ttu-id="eeeaa-132">Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="eeeaa-132">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="eeeaa-133">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-133">**ToolTip**</span></span>    |<span data-ttu-id="eeeaa-134">Нет</span><span class="sxs-lookup"><span data-stu-id="eeeaa-134">No</span></span>|<span data-ttu-id="eeeaa-135">Подсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-135">The tooltip for the button.</span></span> <span data-ttu-id="eeeaa-136">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String.**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="eeeaa-137">**String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="eeeaa-137">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="eeeaa-138">Supertip</span><span class="sxs-lookup"><span data-stu-id="eeeaa-138">Supertip</span></span>](supertip.md)  | <span data-ttu-id="eeeaa-139">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-139">Yes</span></span> |  <span data-ttu-id="eeeaa-140">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-140">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="eeeaa-141">Icon</span><span class="sxs-lookup"><span data-stu-id="eeeaa-141">Icon</span></span>](icon.md)      | <span data-ttu-id="eeeaa-142">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-142">Yes</span></span> |  <span data-ttu-id="eeeaa-143">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-143">An image for the button.</span></span>         |
|  [<span data-ttu-id="eeeaa-144">Action</span><span class="sxs-lookup"><span data-stu-id="eeeaa-144">Action</span></span>](action.md)    | <span data-ttu-id="eeeaa-145">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-145">Yes</span></span> |  <span data-ttu-id="eeeaa-146">Указание действия, которое предстоит выполнить.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-146">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="eeeaa-147">Enabled</span><span class="sxs-lookup"><span data-stu-id="eeeaa-147">Enabled</span></span>](enabled.md)    | <span data-ttu-id="eeeaa-148">Нет</span><span class="sxs-lookup"><span data-stu-id="eeeaa-148">No</span></span> |  <span data-ttu-id="eeeaa-149">Указывает, включен ли этот контроль при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-149">Specifies whether the control is enabled when the add-in launches.</span></span>  |
|  [<span data-ttu-id="eeeaa-150">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="eeeaa-150">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="eeeaa-151">Нет</span><span class="sxs-lookup"><span data-stu-id="eeeaa-151">No</span></span> |  <span data-ttu-id="eeeaa-152">Указывает, должна ли кнопка отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-152">Specifies whether the button should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="eeeaa-153">Если используется, он должен быть *первым элементом.*</span><span class="sxs-lookup"><span data-stu-id="eeeaa-153">If used, it must be the *first* child element.</span></span> |

### <a name="executefunction-button-example"></a><span data-ttu-id="eeeaa-154">Пример кнопки ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="eeeaa-154">ExecuteFunction button example</span></span>

<span data-ttu-id="eeeaa-155">В следующем примере кнопка отключена при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-155">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="eeeaa-156">Его можно включить программным путем.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-156">It can be programmatically enabled.</span></span> <span data-ttu-id="eeeaa-157">Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="eeeaa-157">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
  <Enabled>false</Enabled>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="eeeaa-158">Пример кнопки ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="eeeaa-158">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="eeeaa-159">Элементы управления "Меню" (кнопка с раскрывающимся списком)</span><span class="sxs-lookup"><span data-stu-id="eeeaa-159">Menu (dropdown button) controls</span></span>

<span data-ttu-id="eeeaa-p110">Меню определяет статический список вариантов. Каждый элемент меню либо выполняет функцию, либо отображает область задач. Вложенные меню не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p110">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="eeeaa-163">При использовании с [точкой расширения](extensionpoint.md) **ContextMenu\*\*\*\*PrimaryCommandSurface** элемент управления Menu определяет следующее:</span><span class="sxs-lookup"><span data-stu-id="eeeaa-163">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="eeeaa-164">элемент меню корневого уровня;</span><span class="sxs-lookup"><span data-stu-id="eeeaa-164">A root-level menu item.</span></span>

- <span data-ttu-id="eeeaa-165">список элементов подменю.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-165">A list of submenu items.</span></span>

<span data-ttu-id="eeeaa-p111">При использовании совместно с элементом **PrimaryCommandSurface**, корневой элемент меню отображается в виде кнопки на ленте. При выборе кнопки отображается подменю в виде раскрывающегося списка. При использовании совместно с элементом **ContextMenu**, элемент меню с подменю вставляется в контекстное меню. В обоих случаях индивидуальные элементы подменю могут выполнять функцию JavaScript или отображать область задач. В настоящее время поддерживается только один уровень подменю.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p111">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="eeeaa-p112">В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю отображает область задач, а второй запускает функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p112">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="eeeaa-173">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="eeeaa-173">Child elements</span></span>

|  <span data-ttu-id="eeeaa-174">Элемент</span><span class="sxs-lookup"><span data-stu-id="eeeaa-174">Element</span></span> |  <span data-ttu-id="eeeaa-175">Обязательный</span><span class="sxs-lookup"><span data-stu-id="eeeaa-175">Required</span></span>  |  <span data-ttu-id="eeeaa-176">Описание</span><span class="sxs-lookup"><span data-stu-id="eeeaa-176">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="eeeaa-177">**Label**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-177">**Label**</span></span>     | <span data-ttu-id="eeeaa-178">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-178">Yes</span></span> |  <span data-ttu-id="eeeaa-179">Текст для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-179">The text for the button.</span></span> <span data-ttu-id="eeeaa-180">Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="eeeaa-180">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="eeeaa-181">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-181">**ToolTip**</span></span>    |<span data-ttu-id="eeeaa-182">Нет</span><span class="sxs-lookup"><span data-stu-id="eeeaa-182">No</span></span>|<span data-ttu-id="eeeaa-183">Подсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-183">The tooltip for the button.</span></span> <span data-ttu-id="eeeaa-184">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String.**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-184">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="eeeaa-185">**String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="eeeaa-185">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="eeeaa-186">Supertip</span><span class="sxs-lookup"><span data-stu-id="eeeaa-186">Supertip</span></span>](supertip.md)  | <span data-ttu-id="eeeaa-187">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-187">Yes</span></span> |  <span data-ttu-id="eeeaa-188">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-188">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="eeeaa-189">Icon</span><span class="sxs-lookup"><span data-stu-id="eeeaa-189">Icon</span></span>](icon.md)      | <span data-ttu-id="eeeaa-190">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-190">Yes</span></span> |  <span data-ttu-id="eeeaa-191">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-191">An image for the button.</span></span>         |
|  <span data-ttu-id="eeeaa-192">**Items**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-192">**Items**</span></span>     | <span data-ttu-id="eeeaa-193">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-193">Yes</span></span> |  <span data-ttu-id="eeeaa-194">Коллекция кнопок, отображающихся в меню.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-194">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="eeeaa-195">Содержит элементы **Item** для каждого элемента подменю.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-195">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="eeeaa-196">Каждый элемент **Item** содержит дочерние элементы, вложенные в [элемент управления Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="eeeaa-196">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|
|  [<span data-ttu-id="eeeaa-197">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="eeeaa-197">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="eeeaa-198">Нет</span><span class="sxs-lookup"><span data-stu-id="eeeaa-198">No</span></span> |  <span data-ttu-id="eeeaa-199">Указывает, должно ли меню отображаться в сочетаниях приложений и платформ, которые поддерживают настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-199">Specifies whether the menu should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="eeeaa-200">Если используется, он должен быть *первым элементом.*</span><span class="sxs-lookup"><span data-stu-id="eeeaa-200">If used, it must be the *first* child element.</span></span> |

### <a name="menu-control-examples"></a><span data-ttu-id="eeeaa-201">Примеры элементов управления Menu</span><span class="sxs-lookup"><span data-stu-id="eeeaa-201">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="eeeaa-202">Элемент управления MobileButton</span><span class="sxs-lookup"><span data-stu-id="eeeaa-202">MobileButton control</span></span>

<span data-ttu-id="eeeaa-p117">Кнопка мобильного устройства выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка мобильного устройства" должен иметь атрибут `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p117">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="eeeaa-p118">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-p118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="eeeaa-208">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="eeeaa-208">Child elements</span></span>
|  <span data-ttu-id="eeeaa-209">Элемент</span><span class="sxs-lookup"><span data-stu-id="eeeaa-209">Element</span></span> |  <span data-ttu-id="eeeaa-210">Обязательный</span><span class="sxs-lookup"><span data-stu-id="eeeaa-210">Required</span></span>  |  <span data-ttu-id="eeeaa-211">Описание</span><span class="sxs-lookup"><span data-stu-id="eeeaa-211">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="eeeaa-212">**Label**</span><span class="sxs-lookup"><span data-stu-id="eeeaa-212">**Label**</span></span>     | <span data-ttu-id="eeeaa-213">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-213">Yes</span></span> |  <span data-ttu-id="eeeaa-214">Текст для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-214">The text for the button.</span></span> <span data-ttu-id="eeeaa-215">Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="eeeaa-215">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="eeeaa-216">Icon</span><span class="sxs-lookup"><span data-stu-id="eeeaa-216">Icon</span></span>](icon.md)      | <span data-ttu-id="eeeaa-217">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-217">Yes</span></span> |  <span data-ttu-id="eeeaa-218">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-218">An image for the button.</span></span>         |
|  [<span data-ttu-id="eeeaa-219">Action</span><span class="sxs-lookup"><span data-stu-id="eeeaa-219">Action</span></span>](action.md)    | <span data-ttu-id="eeeaa-220">Да</span><span class="sxs-lookup"><span data-stu-id="eeeaa-220">Yes</span></span> |  <span data-ttu-id="eeeaa-221">Указание действия, которое предстоит выполнить.</span><span class="sxs-lookup"><span data-stu-id="eeeaa-221">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="eeeaa-222">Пример кнопки ExecuteFunction для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="eeeaa-222">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="eeeaa-223">Пример кнопки ShowTaskpane для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="eeeaa-223">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
