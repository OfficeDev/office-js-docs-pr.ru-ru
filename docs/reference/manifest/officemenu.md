# <a name="officemenu-element"></a><span data-ttu-id="a41f2-101">Элемент OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="a41f2-101">OfficeMenu element</span></span>

<span data-ttu-id="a41f2-p101">Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="a41f2-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="a41f2-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a41f2-104">Attributes</span></span>

| <span data-ttu-id="a41f2-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="a41f2-105">Attribute</span></span>            | <span data-ttu-id="a41f2-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a41f2-106">Required</span></span> | <span data-ttu-id="a41f2-107">Описание</span><span class="sxs-lookup"><span data-stu-id="a41f2-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="a41f2-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a41f2-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="a41f2-109">Да</span><span class="sxs-lookup"><span data-stu-id="a41f2-109">Yes</span></span>      | <span data-ttu-id="a41f2-110">Тип определяемого элемента OfficeMenu.</span><span class="sxs-lookup"><span data-stu-id="a41f2-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="a41f2-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="a41f2-111">Child elements</span></span>

|  <span data-ttu-id="a41f2-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="a41f2-112">Element</span></span> |  <span data-ttu-id="a41f2-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a41f2-113">Required</span></span>  |  <span data-ttu-id="a41f2-114">Описание</span><span class="sxs-lookup"><span data-stu-id="a41f2-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a41f2-115">Control</span><span class="sxs-lookup"><span data-stu-id="a41f2-115">Control</span></span>](#control)    | <span data-ttu-id="a41f2-116">Да</span><span class="sxs-lookup"><span data-stu-id="a41f2-116">Yes</span></span> |  <span data-ttu-id="a41f2-117">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="a41f2-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="a41f2-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a41f2-118">xsi:type</span></span>

<span data-ttu-id="a41f2-119">Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="a41f2-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="a41f2-p102">`ContextMenuText` — отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="a41f2-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="a41f2-p103">`ContextMenuCell` — отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel.</span><span class="sxs-lookup"><span data-stu-id="a41f2-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="a41f2-124">Control</span><span class="sxs-lookup"><span data-stu-id="a41f2-124">Control</span></span>

<span data-ttu-id="a41f2-125">Для каждого элемента **OfficeMenu** требуется один или несколько элементов управления [меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="a41f2-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="a41f2-126">Пример</span><span class="sxs-lookup"><span data-stu-id="a41f2-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
