---
title: Элемент OfficeMenu в файле манифеста
description: Элемент OfficeMenu определяет коллекцию элементов управления, добавляемых в контекстное меню Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d181e0c6f489997a149b9713bdc257f4a2baeb16
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641446"
---
# <a name="officemenu-element"></a><span data-ttu-id="20cf7-103">Элемент OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="20cf7-103">OfficeMenu element</span></span>

<span data-ttu-id="20cf7-p101">Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="20cf7-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="20cf7-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="20cf7-106">Attributes</span></span>

| <span data-ttu-id="20cf7-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="20cf7-107">Attribute</span></span>            | <span data-ttu-id="20cf7-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="20cf7-108">Required</span></span> | <span data-ttu-id="20cf7-109">Описание</span><span class="sxs-lookup"><span data-stu-id="20cf7-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="20cf7-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="20cf7-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="20cf7-111">Да</span><span class="sxs-lookup"><span data-stu-id="20cf7-111">Yes</span></span>      | <span data-ttu-id="20cf7-112">Тип определяемого элемента OfficeMenu.</span><span class="sxs-lookup"><span data-stu-id="20cf7-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="20cf7-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="20cf7-113">Child elements</span></span>

|  <span data-ttu-id="20cf7-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="20cf7-114">Element</span></span> |  <span data-ttu-id="20cf7-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="20cf7-115">Required</span></span>  |  <span data-ttu-id="20cf7-116">Описание</span><span class="sxs-lookup"><span data-stu-id="20cf7-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="20cf7-117">Control</span><span class="sxs-lookup"><span data-stu-id="20cf7-117">Control</span></span>](#control)    | <span data-ttu-id="20cf7-118">Да</span><span class="sxs-lookup"><span data-stu-id="20cf7-118">Yes</span></span> |  <span data-ttu-id="20cf7-119">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="20cf7-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="20cf7-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="20cf7-120">xsi:type</span></span>

<span data-ttu-id="20cf7-121">Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="20cf7-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="20cf7-p102">`ContextMenuText`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="20cf7-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="20cf7-p103">`ContextMenuCell`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel.</span><span class="sxs-lookup"><span data-stu-id="20cf7-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span>

## <a name="control"></a><span data-ttu-id="20cf7-126">Control</span><span class="sxs-lookup"><span data-stu-id="20cf7-126">Control</span></span>

<span data-ttu-id="20cf7-127">Для каждого элемента **OfficeMenu** требуется один или несколько элементов управления [меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="20cf7-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="20cf7-128">Пример</span><span class="sxs-lookup"><span data-stu-id="20cf7-128">Example</span></span>

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
