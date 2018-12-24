---
title: Элемент OfficeMenu в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d243612c9b78c362bed9d90dcb539b0dbacfa6f3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432489"
---
# <a name="officemenu-element"></a><span data-ttu-id="7486f-102">Элемент OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="7486f-102">OfficeMenu element</span></span>

<span data-ttu-id="7486f-p101">Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="7486f-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="7486f-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7486f-105">Attributes</span></span>

| <span data-ttu-id="7486f-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="7486f-106">Attribute</span></span>            | <span data-ttu-id="7486f-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="7486f-107">Required</span></span> | <span data-ttu-id="7486f-108">Описание</span><span class="sxs-lookup"><span data-stu-id="7486f-108">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="7486f-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7486f-109">xsi:type</span></span>](#xsitype) | <span data-ttu-id="7486f-110">Да</span><span class="sxs-lookup"><span data-stu-id="7486f-110">Yes</span></span>      | <span data-ttu-id="7486f-111">Тип определяемого элемента OfficeMenu.</span><span class="sxs-lookup"><span data-stu-id="7486f-111">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="7486f-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="7486f-112">Child elements</span></span>

|  <span data-ttu-id="7486f-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="7486f-113">Element</span></span> |  <span data-ttu-id="7486f-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="7486f-114">Required</span></span>  |  <span data-ttu-id="7486f-115">Описание</span><span class="sxs-lookup"><span data-stu-id="7486f-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7486f-116">Control</span><span class="sxs-lookup"><span data-stu-id="7486f-116">Control</span></span>](#control)    | <span data-ttu-id="7486f-117">Да</span><span class="sxs-lookup"><span data-stu-id="7486f-117">Yes</span></span> |  <span data-ttu-id="7486f-118">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="7486f-118">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="7486f-119">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7486f-119">xsi:type</span></span>

<span data-ttu-id="7486f-120">Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="7486f-120">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="7486f-p102">`ContextMenuText`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="7486f-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="7486f-p103">`ContextMenuCell`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel.</span><span class="sxs-lookup"><span data-stu-id="7486f-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="7486f-125">Control</span><span class="sxs-lookup"><span data-stu-id="7486f-125">Control</span></span>

<span data-ttu-id="7486f-126">Для каждого элемента **OfficeMenu** требуется один или несколько элементов управления [меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="7486f-126">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="7486f-127">Пример</span><span class="sxs-lookup"><span data-stu-id="7486f-127">Example</span></span>

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
