---
title: Области задач в надстройках Office
description: Области задач предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или сообщений электронной почты, а также для отображения данных из источника данных.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093758"
---
# <a name="task-panes-in-office-add-ins"></a>Области задач в надстройках Office
 
Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Рис. 1. Типичный макет области задач*

![Типичный макет области задач](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Советы и рекомендации

|**Рекомендуется**|**Не рекомендуется**|
|:-----|:--------|
|<ul><li>Включите имя надстройки в название.</li></ul>|<ul><li>Не включайте в него название вашей компании.</li></ul>|
|<ul><li>Используйте короткие описательные имена в названии.</li></ul>|<ul><li>Не добавляйте такие строки, как "надстройка", "для Word" или "для Office", в название надстройки.</li></ul>|
|<ul><li>Добавьте элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.</li></ul>||
|<ul><li>Включите элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки, если только она не будет использоваться исключительно в Outlook.</li></ul>||


## <a name="variants"></a>Варианты

The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*Рис. 2. Размеры области задач в классических приложениях Office 2016*

![Размеры области задач в классических приложениях при разрешении 1366 x 768](../images/office-2016-taskpane-sizes.png)

- Excel — 320 x 455
- PowerPoint — 320 x 531
- Word — 320 x 531
- Outlook — 348 x 535

<br/>

*Рис. 3. Размеры областей задач Office*

![Размеры области задач в классических приложениях при разрешении 1366 x 768](../images/office-365-taskpane-sizes.png)

- Excel — 350 x 378
- PowerPoint — 348 x 391
- Word — 329 x 445
- Outlook (в Интернете) — 320 x 570

## <a name="personality-menu"></a>Меню личных данных

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

Меню личных данных в Windows имеет размер 12 x 32 пикселей, как показано ниже.

*Рис. 4. Меню личных данных в Windows*

![Меню личных данных на компьютере с Windows](../images/personality-menu-win.png)

В Mac меню личных данных имеет размер 26 x 26 пикселей, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселя, как показано на изображении.

*Рис. 5. Меню личных данных на Mac*

![Меню личных данных на Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Реализация

Ознакомьтесь с реализацией области задач на примере [надстройки Excel "Тенденции расходов банка WoodGrove" на JS](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) на сайте GitHub. 


## <a name="see-also"></a>См. также

- [Office UI Fabric в надстройках Office](office-ui-fabric.md) 
- [Конструктивные шаблоны для надстроек Office](../design/ux-design-pattern-templates.md)

