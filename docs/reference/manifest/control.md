---
title: Элемент Control в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d77b464fde9898ef216ef9e47c651fb5750e4453
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450641"
---
# <a name="control-element"></a><span data-ttu-id="81bea-102">Элемент Control</span><span class="sxs-lookup"><span data-stu-id="81bea-102">Control element</span></span>

<span data-ttu-id="81bea-p101">Определяет функцию JavaScript, которая выполняет действие или открывает область задач. Элемент **Control** может быть кнопкой или пунктом меню. Элемент [Group](group.md) должен содержать по крайней мере один элемент **Control**.</span><span class="sxs-lookup"><span data-stu-id="81bea-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="81bea-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="81bea-106">Attributes</span></span>

|  <span data-ttu-id="81bea-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="81bea-107">Attribute</span></span>  |  <span data-ttu-id="81bea-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="81bea-108">Required</span></span>  |  <span data-ttu-id="81bea-109">Описание</span><span class="sxs-lookup"><span data-stu-id="81bea-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="81bea-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="81bea-110">**xsi:type**</span></span>|<span data-ttu-id="81bea-111">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-111">Yes</span></span>|<span data-ttu-id="81bea-p102">Тип определяемого элемента управления. Доступные варианты: `Button`, `Menu` или `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="81bea-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="81bea-114">**id**</span><span class="sxs-lookup"><span data-stu-id="81bea-114">**id**</span></span>|<span data-ttu-id="81bea-115">Нет</span><span class="sxs-lookup"><span data-stu-id="81bea-115">No</span></span>|<span data-ttu-id="81bea-p103">ИД элемента управления. Может содержать до 125 знаков.</span><span class="sxs-lookup"><span data-stu-id="81bea-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="81bea-118">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="81bea-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="81bea-119">Применяется только к элементам **Control**, которые содержатся в элементе [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="81bea-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="81bea-120">Элемент управления ''Кнопка''</span><span class="sxs-lookup"><span data-stu-id="81bea-120">Button control</span></span>

<span data-ttu-id="81bea-p105">Кнопка выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка" должен иметь элемент `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="81bea-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="81bea-124">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="81bea-124">Child elements</span></span>
|  <span data-ttu-id="81bea-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="81bea-125">Element</span></span> |  <span data-ttu-id="81bea-126">Обязательный</span><span class="sxs-lookup"><span data-stu-id="81bea-126">Required</span></span>  |  <span data-ttu-id="81bea-127">Описание</span><span class="sxs-lookup"><span data-stu-id="81bea-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="81bea-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="81bea-128">**Label**</span></span>     | <span data-ttu-id="81bea-129">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-129">Yes</span></span> |  <span data-ttu-id="81bea-p106">Текст для кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81bea-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="81bea-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="81bea-132">**ToolTip**</span></span>  |<span data-ttu-id="81bea-133">Нет</span><span class="sxs-lookup"><span data-stu-id="81bea-133">No</span></span>|<span data-ttu-id="81bea-p107">Подсказка для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81bea-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="81bea-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="81bea-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="81bea-138">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-138">Yes</span></span> |  <span data-ttu-id="81bea-139">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="81bea-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="81bea-140">Icon</span><span class="sxs-lookup"><span data-stu-id="81bea-140">Icon</span></span>](icon.md)      | <span data-ttu-id="81bea-141">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-141">Yes</span></span> |  <span data-ttu-id="81bea-142">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="81bea-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="81bea-143">Action</span><span class="sxs-lookup"><span data-stu-id="81bea-143">Action</span></span>](action.md)    | <span data-ttu-id="81bea-144">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-144">Yes</span></span> |  <span data-ttu-id="81bea-145">Указание действия, которое предстоит выполнить.</span><span class="sxs-lookup"><span data-stu-id="81bea-145">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="81bea-146">Пример кнопки ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="81bea-146">ExecuteFunction button example</span></span>

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
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="81bea-147">Пример кнопки ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="81bea-147">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="81bea-148">Элементы управления "Меню" (кнопка с раскрывающимся списком)</span><span class="sxs-lookup"><span data-stu-id="81bea-148">Menu (dropdown button) controls</span></span>

<span data-ttu-id="81bea-p108">Меню определяет статический список вариантов. Каждый элемент меню либо выполняет функцию, либо отображает область задач. Вложенные меню не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="81bea-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="81bea-152">При использовании с [точкой расширения](extensionpoint.md) **ContextMenu\*\*\*\*PrimaryCommandSurface** элемент управления Menu определяет следующее:</span><span class="sxs-lookup"><span data-stu-id="81bea-152">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="81bea-153">элемент меню корневого уровня;</span><span class="sxs-lookup"><span data-stu-id="81bea-153">A root-level menu item.</span></span>

- <span data-ttu-id="81bea-154">список элементов подменю.</span><span class="sxs-lookup"><span data-stu-id="81bea-154">A list of submenu items.</span></span>

<span data-ttu-id="81bea-p109">При использовании с **PrimaryCommandSurface** корневой элемент меню отображает кнопку на ленте. По нажатию кнопки в подменю отображается раскрывающийся список. При использовании с **ContextMenu** в контекстное меню вставляется элемент меню с подменю. В обоих случаях отдельные элементы подменю могут либо вызывать функцию JavaScript, либо отображать область задач. В настоящее время поддерживается только один уровень подменю.</span><span class="sxs-lookup"><span data-stu-id="81bea-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="81bea-p110">В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю отображает область задач, а второй запускает функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="81bea-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="81bea-162">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="81bea-162">Child elements</span></span>

|  <span data-ttu-id="81bea-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="81bea-163">Element</span></span> |  <span data-ttu-id="81bea-164">Обязательный</span><span class="sxs-lookup"><span data-stu-id="81bea-164">Required</span></span>  |  <span data-ttu-id="81bea-165">Описание</span><span class="sxs-lookup"><span data-stu-id="81bea-165">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="81bea-166">**Label**</span><span class="sxs-lookup"><span data-stu-id="81bea-166">**Label**</span></span>     | <span data-ttu-id="81bea-167">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-167">Yes</span></span> |  <span data-ttu-id="81bea-p111">Текст для кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81bea-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="81bea-170">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="81bea-170">**ToolTip**</span></span>  |<span data-ttu-id="81bea-171">Нет</span><span class="sxs-lookup"><span data-stu-id="81bea-171">No</span></span>|<span data-ttu-id="81bea-p112">Подсказка для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81bea-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="81bea-175">Supertip</span><span class="sxs-lookup"><span data-stu-id="81bea-175">Supertip</span></span>](supertip.md)  | <span data-ttu-id="81bea-176">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-176">Yes</span></span> |  <span data-ttu-id="81bea-177">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="81bea-177">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="81bea-178">Icon</span><span class="sxs-lookup"><span data-stu-id="81bea-178">Icon</span></span>](icon.md)      | <span data-ttu-id="81bea-179">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-179">Yes</span></span> |  <span data-ttu-id="81bea-180">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="81bea-180">An image for the button.</span></span>         |
|  <span data-ttu-id="81bea-181">**Items**</span><span class="sxs-lookup"><span data-stu-id="81bea-181">**Items**</span></span>     | <span data-ttu-id="81bea-182">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-182">Yes</span></span> |  <span data-ttu-id="81bea-p113">Коллекция кнопок, отображающихся в меню. Содержит элементы **Item** для каждого элемента подменю. Каждый элемент **Item** содержит дочерние элементы, вложенные в [элемент управления Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="81bea-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="81bea-186">Примеры элементов управления Menu</span><span class="sxs-lookup"><span data-stu-id="81bea-186">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="81bea-187">Элемент управления MobileButton</span><span class="sxs-lookup"><span data-stu-id="81bea-187">MobileButton control</span></span>

<span data-ttu-id="81bea-p114">Кнопка мобильного устройства выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка мобильного устройства" должен иметь атрибут `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="81bea-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="81bea-p115">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="81bea-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="81bea-193">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="81bea-193">Child elements</span></span>
|  <span data-ttu-id="81bea-194">Элемент</span><span class="sxs-lookup"><span data-stu-id="81bea-194">Element</span></span> |  <span data-ttu-id="81bea-195">Обязательный</span><span class="sxs-lookup"><span data-stu-id="81bea-195">Required</span></span>  |  <span data-ttu-id="81bea-196">Описание</span><span class="sxs-lookup"><span data-stu-id="81bea-196">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="81bea-197">**Label**</span><span class="sxs-lookup"><span data-stu-id="81bea-197">**Label**</span></span>     | <span data-ttu-id="81bea-198">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-198">Yes</span></span> |  <span data-ttu-id="81bea-p116">Текст для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, который принадлежит элементу **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81bea-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="81bea-201">Icon</span><span class="sxs-lookup"><span data-stu-id="81bea-201">Icon</span></span>](icon.md)      | <span data-ttu-id="81bea-202">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-202">Yes</span></span> |  <span data-ttu-id="81bea-203">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="81bea-203">An image for the button.</span></span>         |
|  [<span data-ttu-id="81bea-204">Action</span><span class="sxs-lookup"><span data-stu-id="81bea-204">Action</span></span>](action.md)    | <span data-ttu-id="81bea-205">Да</span><span class="sxs-lookup"><span data-stu-id="81bea-205">Yes</span></span> |  <span data-ttu-id="81bea-206">Указание действия, которое предстоит выполнить.</span><span class="sxs-lookup"><span data-stu-id="81bea-206">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="81bea-207">Пример кнопки ExecuteFunction для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="81bea-207">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="81bea-208">Пример кнопки ShowTaskpane для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="81bea-208">ShowTaskpane mobile button example</span></span>

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
