---
title: Элемент OfficeMenu в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 20d020b8ab826049ef0271cbdb8d51201ee88ab4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452020"
---
# <a name="officemenu-element"></a><span data-ttu-id="641e7-102">Элемент OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="641e7-102">OfficeMenu element</span></span>

<span data-ttu-id="641e7-p101">Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="641e7-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="641e7-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="641e7-105">Attributes</span></span>

| <span data-ttu-id="641e7-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="641e7-106">Attribute</span></span>            | <span data-ttu-id="641e7-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="641e7-107">Required</span></span> | <span data-ttu-id="641e7-108">Описание</span><span class="sxs-lookup"><span data-stu-id="641e7-108">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="641e7-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="641e7-109">xsi:type</span></span>](#xsitype) | <span data-ttu-id="641e7-110">Да</span><span class="sxs-lookup"><span data-stu-id="641e7-110">Yes</span></span>      | <span data-ttu-id="641e7-111">Тип определяемого элемента OfficeMenu.</span><span class="sxs-lookup"><span data-stu-id="641e7-111">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="641e7-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="641e7-112">Child elements</span></span>

|  <span data-ttu-id="641e7-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="641e7-113">Element</span></span> |  <span data-ttu-id="641e7-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="641e7-114">Required</span></span>  |  <span data-ttu-id="641e7-115">Описание</span><span class="sxs-lookup"><span data-stu-id="641e7-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="641e7-116">Control</span><span class="sxs-lookup"><span data-stu-id="641e7-116">Control</span></span>](#control)    | <span data-ttu-id="641e7-117">Да</span><span class="sxs-lookup"><span data-stu-id="641e7-117">Yes</span></span> |  <span data-ttu-id="641e7-118">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="641e7-118">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="641e7-119">xsi:type</span><span class="sxs-lookup"><span data-stu-id="641e7-119">xsi:type</span></span>

<span data-ttu-id="641e7-120">Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="641e7-120">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="641e7-p102">`ContextMenuText`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="641e7-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="641e7-p103">`ContextMenuCell`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel.</span><span class="sxs-lookup"><span data-stu-id="641e7-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="641e7-125">Control</span><span class="sxs-lookup"><span data-stu-id="641e7-125">Control</span></span>

<span data-ttu-id="641e7-126">Для каждого элемента **OfficeMenu** требуется один или несколько элементов управления [меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="641e7-126">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="641e7-127">Пример</span><span class="sxs-lookup"><span data-stu-id="641e7-127">Example</span></span>

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
