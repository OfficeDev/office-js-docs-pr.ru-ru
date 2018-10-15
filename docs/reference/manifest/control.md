# <a name="control-element"></a><span data-ttu-id="e0ac0-101">Элемент Control</span><span class="sxs-lookup"><span data-stu-id="e0ac0-101">Control element</span></span>

<span data-ttu-id="e0ac0-p101">Определяет функцию JavaScript, которая выполняет действие или открывает область задач. Элемент **Control** может быть кнопкой или пунктом меню. По крайней мере один элемент **Control** должен быть в элементе [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="e0ac0-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e0ac0-105">Attributes</span></span>

|  <span data-ttu-id="e0ac0-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e0ac0-106">Attribute</span></span>  |  <span data-ttu-id="e0ac0-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e0ac0-107">Required</span></span>  |  <span data-ttu-id="e0ac0-108">Описание</span><span class="sxs-lookup"><span data-stu-id="e0ac0-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="e0ac0-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-109">**xsi:type**</span></span>|<span data-ttu-id="e0ac0-110">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-110">Yes</span></span>|<span data-ttu-id="e0ac0-p102">Тип определяемого элемента управления. Доступные варианты: `Button`, `Menu` или `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="e0ac0-113">**id**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-113">**id**</span></span>|<span data-ttu-id="e0ac0-114">Нет</span><span class="sxs-lookup"><span data-stu-id="e0ac0-114">No</span></span>|<span data-ttu-id="e0ac0-p103">ИД элемента управления. Может содержать до 125 знаков.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="e0ac0-117">Значение  `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing VersionOverrides element must have an  attribute value of .</span></span> <span data-ttu-id="e0ac0-118">Применяется только к элементам **Control**, которые содержатся в элементе [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-118">Note: The  value for xsi:type is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="e0ac0-119">Элемент управления "Кнопка"</span><span class="sxs-lookup"><span data-stu-id="e0ac0-119">Button control</span></span>

<span data-ttu-id="e0ac0-p105">Кнопка выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка" должен иметь элемент `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="e0ac0-123">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e0ac0-123">Child elements</span></span>
|  <span data-ttu-id="e0ac0-124">Элемент</span><span class="sxs-lookup"><span data-stu-id="e0ac0-124">Element</span></span> |  <span data-ttu-id="e0ac0-125">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e0ac0-125">Required</span></span>  |  <span data-ttu-id="e0ac0-126">Описание</span><span class="sxs-lookup"><span data-stu-id="e0ac0-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e0ac0-127">**Label**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-127">**Label**</span></span>     | <span data-ttu-id="e0ac0-128">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-128">Yes</span></span> |  <span data-ttu-id="e0ac0-p106">Текст для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, который принадлежит элементу **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="e0ac0-131">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-131">**ToolTip**</span></span>  |<span data-ttu-id="e0ac0-132">Нет</span><span class="sxs-lookup"><span data-stu-id="e0ac0-132">No</span></span>|<span data-ttu-id="e0ac0-p107">Подсказка для кнопки. Для атрибута **resid**  должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String**  — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="e0ac0-136">Supertip</span><span class="sxs-lookup"><span data-stu-id="e0ac0-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="e0ac0-137">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-137">Yes</span></span> |  <span data-ttu-id="e0ac0-138">Суперподсказка для кнопки.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="e0ac0-139">Icon</span><span class="sxs-lookup"><span data-stu-id="e0ac0-139">Icon</span></span>](icon.md)      | <span data-ttu-id="e0ac0-140">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-140">Yes</span></span> |  <span data-ttu-id="e0ac0-141">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="e0ac0-142">Action</span><span class="sxs-lookup"><span data-stu-id="e0ac0-142">Action</span></span>](action.md)    | <span data-ttu-id="e0ac0-143">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-143">Yes</span></span> |  <span data-ttu-id="e0ac0-144">Задает выполняемое действие.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="e0ac0-145">Пример кнопки ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="e0ac0-145">ExecuteFunction button example</span></span>

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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="e0ac0-146">Пример кнопки ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="e0ac0-146">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="e0ac0-147">Элементы управления "Меню" (кнопка с раскрывающимся списком)</span><span class="sxs-lookup"><span data-stu-id="e0ac0-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="e0ac0-p108">Меню определяет статический список вариантов. Каждый элемент меню либо выполняет функцию, либо отображает область задач. Вложенные меню не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="e0ac0-151">При исользовании с **PrimaryCommandSurface** или **ContextMenu** [точками расширения](extensionpoint.md), элемент управления меню определяет:</span><span class="sxs-lookup"><span data-stu-id="e0ac0-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="e0ac0-152">элемент меню корневого уровня;</span><span class="sxs-lookup"><span data-stu-id="e0ac0-152">A root-level menu item.</span></span>

- <span data-ttu-id="e0ac0-153">список элементов подменю.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-153">A list of submenu items.</span></span>

<span data-ttu-id="e0ac0-p109">При использовании с **PrimaryCommandSurface** корневой элемент меню отображает кнопку на ленте. По нажатию кнопки в подменю отображается раскрывающийся список. При использовании с **ContextMenu** в контекстное меню вставляется элемент меню с подменю. В обоих случаях отдельные элементы подменю могут либо вызывать функцию JavaScript, либо отображать область задач. В настоящее время поддерживается только один уровень подменю.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="e0ac0-p110">В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю отображает область задач, а второй запускает функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="e0ac0-161">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e0ac0-161">Child elements</span></span>

|  <span data-ttu-id="e0ac0-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="e0ac0-162">Element</span></span> |  <span data-ttu-id="e0ac0-163">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e0ac0-163">Required</span></span>  |  <span data-ttu-id="e0ac0-164">Описание</span><span class="sxs-lookup"><span data-stu-id="e0ac0-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e0ac0-165">**Label**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-165">**Label**</span></span>     | <span data-ttu-id="e0ac0-166">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-166">Yes</span></span> |  <span data-ttu-id="e0ac0-p111">Текст для кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="e0ac0-169">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-169">**ToolTip**</span></span>  |<span data-ttu-id="e0ac0-170">Нет</span><span class="sxs-lookup"><span data-stu-id="e0ac0-170">No</span></span>|<span data-ttu-id="e0ac0-p112">Подсказка для кнопки. Для атрибута **resid**  должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String**  — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="e0ac0-174">Supertip</span><span class="sxs-lookup"><span data-stu-id="e0ac0-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="e0ac0-175">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-175">Yes</span></span> |  <span data-ttu-id="e0ac0-176">Суперподсказка для этой кнопки.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="e0ac0-177">Icon</span><span class="sxs-lookup"><span data-stu-id="e0ac0-177">Icon</span></span>](icon.md)      | <span data-ttu-id="e0ac0-178">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-178">Yes</span></span> |  <span data-ttu-id="e0ac0-179">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-179">An image for the button.</span></span>         |
|  <span data-ttu-id="e0ac0-180">**Элементы**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-180">**Items**</span></span>     | <span data-ttu-id="e0ac0-181">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-181">Yes</span></span> |  <span data-ttu-id="e0ac0-p113">Коллекция кнопок, отображающихся в меню. Содержит элементы **Item** для каждого элемента подменю. Каждый элемент **Item** содержит дочерние элементы, вложенные в [элемент управления Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="e0ac0-185">Примеры элементов управления "Меню"</span><span class="sxs-lookup"><span data-stu-id="e0ac0-185">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="e0ac0-186">Элемент управления MobileButton</span><span class="sxs-lookup"><span data-stu-id="e0ac0-186">MobileButton control</span></span>

<span data-ttu-id="e0ac0-p114">Кнопка мобильного устройства выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка мобильного устройства" должен иметь атрибут `id`, уникальный для манифеста.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="e0ac0-p115">Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение  атрибута `xsi:type` `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="e0ac0-192">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e0ac0-192">Child elements</span></span>
|  <span data-ttu-id="e0ac0-193">Элемент</span><span class="sxs-lookup"><span data-stu-id="e0ac0-193">Element</span></span> |  <span data-ttu-id="e0ac0-194">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e0ac0-194">Required</span></span>  |  <span data-ttu-id="e0ac0-195">Описание</span><span class="sxs-lookup"><span data-stu-id="e0ac0-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e0ac0-196">**Label**</span><span class="sxs-lookup"><span data-stu-id="e0ac0-196">**Label**</span></span>     | <span data-ttu-id="e0ac0-197">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-197">Yes</span></span> |  <span data-ttu-id="e0ac0-p116">Текст для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, который принадлежит элементу **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e0ac0-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="e0ac0-200">Icon</span><span class="sxs-lookup"><span data-stu-id="e0ac0-200">Icon</span></span>](icon.md)      | <span data-ttu-id="e0ac0-201">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-201">Yes</span></span> |  <span data-ttu-id="e0ac0-202">Изображение для кнопки.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="e0ac0-203">Action</span><span class="sxs-lookup"><span data-stu-id="e0ac0-203">Action</span></span>](action.md)    | <span data-ttu-id="e0ac0-204">Да</span><span class="sxs-lookup"><span data-stu-id="e0ac0-204">Yes</span></span> |  <span data-ttu-id="e0ac0-205">Задает выполняемое действие.</span><span class="sxs-lookup"><span data-stu-id="e0ac0-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="e0ac0-206">Пример кнопки ExecuteFunction для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="e0ac0-206">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="e0ac0-207">Пример кнопки ShowTaskpane для мобильного устройства</span><span class="sxs-lookup"><span data-stu-id="e0ac0-207">ShowTaskpane mobile button example</span></span>

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