---
title: Элемент Control в файле манифеста
description: ''
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 0add0d102b62411b67c081b74ecd0a138df3b625
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596475"
---
# <a name="control-element"></a><span data-ttu-id="c43c6-102">Элемент Control</span><span class="sxs-lookup"><span data-stu-id="c43c6-102">Control element</span></span>

<span data-ttu-id="c43c6-p101">Определяет функцию JavaScript, которая выполняет действие или открывает область задач. Элемент **Control** может быть кнопкой или пунктом меню. Элемент [Group](group.md) должен содержать по крайней мере один элемент **Control**.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="c43c6-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c43c6-106">Attributes</span></span>

|  <span data-ttu-id="c43c6-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c43c6-107">Attribute</span></span>  |  <span data-ttu-id="c43c6-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c43c6-108">Required</span></span>  |  <span data-ttu-id="c43c6-109">Описание</span><span class="sxs-lookup"><span data-stu-id="c43c6-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="c43c6-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c43c6-110">**xsi:type**</span></span>|<span data-ttu-id="c43c6-111">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-111">Yes</span></span>|<span data-ttu-id="c43c6-p102">Тип определяемого элемента управления. Доступные варианты: `Button`, `Menu` или `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="c43c6-114">**id**</span><span class="sxs-lookup"><span data-stu-id="c43c6-114">**id**</span></span>|<span data-ttu-id="c43c6-115">Нет</span><span class="sxs-lookup"><span data-stu-id="c43c6-115">No</span></span>|<span data-ttu-id="c43c6-p103">ИД элемента управления. Может содержать до 125 знаков.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="c43c6-118">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="c43c6-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="c43c6-119">Применяется только к элементам **Control**, которые содержатся в элементе [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="c43c6-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="c43c6-120">Элемент управления ''Кнопка''</span><span class="sxs-lookup"><span data-stu-id="c43c6-120">Button control</span></span>

<span data-ttu-id="c43c6-p105">Кнопка выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка" должен иметь элемент `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="c43c6-124">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c43c6-124">Child elements</span></span>
|  <span data-ttu-id="c43c6-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="c43c6-125">Element</span></span> |  <span data-ttu-id="c43c6-126">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c43c6-126">Required</span></span>  |  <span data-ttu-id="c43c6-127">Описание</span><span class="sxs-lookup"><span data-stu-id="c43c6-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c43c6-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="c43c6-128">**Label**</span></span>     | <span data-ttu-id="c43c6-129">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-129">Yes</span></span> |  <span data-ttu-id="c43c6-130">Текст для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-130">The text for the button.</span></span> <span data-ttu-id="c43c6-131">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c43c6-131">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="c43c6-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="c43c6-132">**ToolTip**</span></span>    |<span data-ttu-id="c43c6-133">Нет</span><span class="sxs-lookup"><span data-stu-id="c43c6-133">No</span></span>|<span data-ttu-id="c43c6-134">Всплывающая подсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-134">The tooltip for the button.</span></span> <span data-ttu-id="c43c6-135">Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**.</span><span class="sxs-lookup"><span data-stu-id="c43c6-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c43c6-136">**String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c43c6-136">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="c43c6-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="c43c6-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="c43c6-138">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-138">Yes</span></span> |  <span data-ttu-id="c43c6-139">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="c43c6-140">Icon</span><span class="sxs-lookup"><span data-stu-id="c43c6-140">Icon</span></span>](icon.md)      | <span data-ttu-id="c43c6-141">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-141">Yes</span></span> |  <span data-ttu-id="c43c6-142">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="c43c6-143">Action</span><span class="sxs-lookup"><span data-stu-id="c43c6-143">Action</span></span>](action.md)    | <span data-ttu-id="c43c6-144">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-144">Yes</span></span> |  <span data-ttu-id="c43c6-145">Указание действия, которое предстоит выполнить.</span><span class="sxs-lookup"><span data-stu-id="c43c6-145">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="c43c6-146">Enabled</span><span class="sxs-lookup"><span data-stu-id="c43c6-146">Enabled</span></span>](enabled.md)    | <span data-ttu-id="c43c6-147">Нет</span><span class="sxs-lookup"><span data-stu-id="c43c6-147">No</span></span> |  <span data-ttu-id="c43c6-148">Указывает, включен ли элемент управления при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-148">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="c43c6-149">Пример кнопки ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="c43c6-149">ExecuteFunction button example</span></span>

<span data-ttu-id="c43c6-150">В следующем примере кнопка отключается при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-150">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="c43c6-151">Его можно включить программным способом.</span><span class="sxs-lookup"><span data-stu-id="c43c6-151">It can be programmatically enabled.</span></span> <span data-ttu-id="c43c6-152">Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="c43c6-152">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="c43c6-153">Пример кнопки ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="c43c6-153">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="c43c6-154">Элементы управления "Меню" (кнопка с раскрывающимся списком)</span><span class="sxs-lookup"><span data-stu-id="c43c6-154">Menu (dropdown button) controls</span></span>

<span data-ttu-id="c43c6-p109">Меню определяет статический список вариантов. Каждый элемент меню либо выполняет функцию, либо отображает область задач. Вложенные меню не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="c43c6-158">При использовании с [точкой расширения](extensionpoint.md) **ContextMenu\*\*\*\*PrimaryCommandSurface** элемент управления Menu определяет следующее:</span><span class="sxs-lookup"><span data-stu-id="c43c6-158">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="c43c6-159">элемент меню корневого уровня;</span><span class="sxs-lookup"><span data-stu-id="c43c6-159">A root-level menu item.</span></span>

- <span data-ttu-id="c43c6-160">список элементов подменю.</span><span class="sxs-lookup"><span data-stu-id="c43c6-160">A list of submenu items.</span></span>

<span data-ttu-id="c43c6-p110">При использовании совместно с элементом **PrimaryCommandSurface**, корневой элемент меню отображается в виде кнопки на ленте. При выборе кнопки отображается подменю в виде раскрывающегося списка. При использовании совместно с элементом **ContextMenu**, элемент меню с подменю вставляется в контекстное меню. В обоих случаях индивидуальные элементы подменю могут выполнять функцию JavaScript или отображать область задач. В настоящее время поддерживается только один уровень подменю.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="c43c6-p111">В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю отображает область задач, а второй запускает функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="c43c6-168">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c43c6-168">Child elements</span></span>

|  <span data-ttu-id="c43c6-169">Элемент</span><span class="sxs-lookup"><span data-stu-id="c43c6-169">Element</span></span> |  <span data-ttu-id="c43c6-170">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c43c6-170">Required</span></span>  |  <span data-ttu-id="c43c6-171">Описание</span><span class="sxs-lookup"><span data-stu-id="c43c6-171">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c43c6-172">**Label**</span><span class="sxs-lookup"><span data-stu-id="c43c6-172">**Label**</span></span>     | <span data-ttu-id="c43c6-173">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-173">Yes</span></span> |  <span data-ttu-id="c43c6-174">Текст для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-174">The text for the button.</span></span> <span data-ttu-id="c43c6-175">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c43c6-175">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="c43c6-176">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="c43c6-176">**ToolTip**</span></span>    |<span data-ttu-id="c43c6-177">Нет</span><span class="sxs-lookup"><span data-stu-id="c43c6-177">No</span></span>|<span data-ttu-id="c43c6-178">Всплывающая подсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-178">The tooltip for the button.</span></span> <span data-ttu-id="c43c6-179">Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**.</span><span class="sxs-lookup"><span data-stu-id="c43c6-179">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c43c6-180">**String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c43c6-180">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="c43c6-181">Supertip</span><span class="sxs-lookup"><span data-stu-id="c43c6-181">Supertip</span></span>](supertip.md)  | <span data-ttu-id="c43c6-182">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-182">Yes</span></span> |  <span data-ttu-id="c43c6-183">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-183">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="c43c6-184">Icon</span><span class="sxs-lookup"><span data-stu-id="c43c6-184">Icon</span></span>](icon.md)      | <span data-ttu-id="c43c6-185">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-185">Yes</span></span> |  <span data-ttu-id="c43c6-186">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-186">An image for the button.</span></span>         |
|  <span data-ttu-id="c43c6-187">**Items**</span><span class="sxs-lookup"><span data-stu-id="c43c6-187">**Items**</span></span>     | <span data-ttu-id="c43c6-188">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-188">Yes</span></span> |  <span data-ttu-id="c43c6-189">Коллекция кнопок, отображающихся в меню.</span><span class="sxs-lookup"><span data-stu-id="c43c6-189">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="c43c6-190">Содержит элементы **Item** для каждого элемента подменю.</span><span class="sxs-lookup"><span data-stu-id="c43c6-190">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="c43c6-191">Каждый элемент **Item** содержит дочерние элементы, вложенные в [элемент управления Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="c43c6-191">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="c43c6-192">Примеры элементов управления Menu</span><span class="sxs-lookup"><span data-stu-id="c43c6-192">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="c43c6-193">Элемент управления MobileButton</span><span class="sxs-lookup"><span data-stu-id="c43c6-193">MobileButton control</span></span>

<span data-ttu-id="c43c6-p115">Кнопка мобильного устройства выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка мобильного устройства" должен иметь атрибут `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="c43c6-p116">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="c43c6-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="c43c6-199">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c43c6-199">Child elements</span></span>
|  <span data-ttu-id="c43c6-200">Элемент</span><span class="sxs-lookup"><span data-stu-id="c43c6-200">Element</span></span> |  <span data-ttu-id="c43c6-201">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c43c6-201">Required</span></span>  |  <span data-ttu-id="c43c6-202">Описание</span><span class="sxs-lookup"><span data-stu-id="c43c6-202">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c43c6-203">**Label**</span><span class="sxs-lookup"><span data-stu-id="c43c6-203">**Label**</span></span>     | <span data-ttu-id="c43c6-204">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-204">Yes</span></span> |  <span data-ttu-id="c43c6-205">Текст для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-205">The text for the button.</span></span> <span data-ttu-id="c43c6-206">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c43c6-206">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="c43c6-207">Icon</span><span class="sxs-lookup"><span data-stu-id="c43c6-207">Icon</span></span>](icon.md)      | <span data-ttu-id="c43c6-208">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-208">Yes</span></span> |  <span data-ttu-id="c43c6-209">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="c43c6-209">An image for the button.</span></span>         |
|  [<span data-ttu-id="c43c6-210">Action</span><span class="sxs-lookup"><span data-stu-id="c43c6-210">Action</span></span>](action.md)    | <span data-ttu-id="c43c6-211">Да</span><span class="sxs-lookup"><span data-stu-id="c43c6-211">Yes</span></span> |  <span data-ttu-id="c43c6-212">Указание действия, которое предстоит выполнить.</span><span class="sxs-lookup"><span data-stu-id="c43c6-212">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="c43c6-213">Пример кнопки ExecuteFunction для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="c43c6-213">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="c43c6-214">Пример кнопки ShowTaskpane для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="c43c6-214">ShowTaskpane mobile button example</span></span>

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
