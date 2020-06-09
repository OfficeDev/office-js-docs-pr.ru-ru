---
title: Элемент OfficeMenu в файле манифеста
description: Элемент OfficeMenu определяет коллекцию элементов управления, добавляемых в контекстное меню Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f5aac4e3454e1aa18021c10bfb2f06df90805980
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611521"
---
# <a name="officemenu-element"></a><span data-ttu-id="d79ef-103">Элемент OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="d79ef-103">OfficeMenu element</span></span>

<span data-ttu-id="d79ef-p101">Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="d79ef-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="d79ef-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d79ef-106">Attributes</span></span>

| <span data-ttu-id="d79ef-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="d79ef-107">Attribute</span></span>            | <span data-ttu-id="d79ef-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d79ef-108">Required</span></span> | <span data-ttu-id="d79ef-109">Описание</span><span class="sxs-lookup"><span data-stu-id="d79ef-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="d79ef-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d79ef-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="d79ef-111">Да</span><span class="sxs-lookup"><span data-stu-id="d79ef-111">Yes</span></span>      | <span data-ttu-id="d79ef-112">Тип определяемого элемента OfficeMenu.</span><span class="sxs-lookup"><span data-stu-id="d79ef-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="d79ef-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d79ef-113">Child elements</span></span>

|  <span data-ttu-id="d79ef-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="d79ef-114">Element</span></span> |  <span data-ttu-id="d79ef-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d79ef-115">Required</span></span>  |  <span data-ttu-id="d79ef-116">Описание</span><span class="sxs-lookup"><span data-stu-id="d79ef-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d79ef-117">Control</span><span class="sxs-lookup"><span data-stu-id="d79ef-117">Control</span></span>](#control)    | <span data-ttu-id="d79ef-118">Да</span><span class="sxs-lookup"><span data-stu-id="d79ef-118">Yes</span></span> |  <span data-ttu-id="d79ef-119">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="d79ef-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="d79ef-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d79ef-120">xsi:type</span></span>

<span data-ttu-id="d79ef-121">Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="d79ef-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="d79ef-p102">`ContextMenuText`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="d79ef-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="d79ef-p103">`ContextMenuCell`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel.</span><span class="sxs-lookup"><span data-stu-id="d79ef-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="d79ef-126">Control</span><span class="sxs-lookup"><span data-stu-id="d79ef-126">Control</span></span>

<span data-ttu-id="d79ef-127">Для каждого элемента **OfficeMenu** требуется один или несколько элементов управления [меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="d79ef-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="d79ef-128">Пример</span><span class="sxs-lookup"><span data-stu-id="d79ef-128">Example</span></span>

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
