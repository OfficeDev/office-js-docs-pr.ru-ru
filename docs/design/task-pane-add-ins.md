---
title: Области задач в надстройках Office
description: Области задач предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или сообщений электронной почты, а также для отображения данных из источника данных.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: d911101a7df1f1ad8aa01b8e0006bd93d994a193
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092919"
---
# <a name="task-panes-in-office-add-ins"></a>Области задач в надстройках Office

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Рис. 1. Типичный макет области задач*

![Иллюстрация, на которой показан типичный макет области задач с вкладками разделов в верхней части, логотипом компании и названием компании в левом нижнем углу и значком параметров в правом нижнем углу.](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Лучшие методики

|Правильно|Неправильно|
|:-----|:--------|
|Включите имя надстройки в название.|Не включайте в него название вашей компании.|
|Используйте короткие описательные имена в названии.|Не добавляйте строки, такие как "надстройка", "для Word" или "для Office", к заголовку надстройки.|
|Добавьте элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.|*Ни один.*|
|Включите элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки, если только она не будет использоваться исключительно в Outlook.|*Ни один.*|

## <a name="variants"></a>Варианты

На следующих изображениях показаны различные размеры области задач с лентой приложения Office в разрешении 1366x768. Чтобы вставить строку формул в Excel, требуется дополнительное пространство по вертикали.  

*Рис. 2. Размеры области задач в классических приложениях Office 2016*

![Схема размеров области задач рабочего стола с разрешением 1366x768.](../images/office-2016-taskpane-sizes.png)

- Excel — 320 x 455 пикселей
- PowerPoint — 320 x 531 пиксель
- Word — 320 x 531 пиксель
- Outlook — 348 x 535 пикселей

<br/>

*Рис. 3. Размеры области задач Office*

![Схема размеров области задач с разрешением 1366 x 768.](../images/office-365-taskpane-sizes.png)

- Excel — 350 x 378 пикселей
- PowerPoint — 348 x 391 пиксель
- Word — 329 x 445 пикселей
- Outlook (в Интернете) — 320 x 570 пикселей

## <a name="personality-menu"></a>Меню личных данных

Меню личных данных могут перекрывать элементы навигации и управления, расположенные в правой верхней части надстройки. Ниже указаны текущие размеры меню личных данных в Windows и Mac. (Меню личных данных не поддерживается в Outlook.)

Меню личных данных в Windows имеет размер 12 x 32 пикселей, как показано ниже.

*Рис. 4. Меню личных данных в Windows*

![Схема, показывающая меню личных данных на рабочем столе Windows.](../images/personality-menu-win.png)

В Mac меню личных данных имеет размер 26 x 26 пикселей, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселя, как показано на изображении.

*Рис. 5. Меню личных данных на Mac*

![Схема, показывающая меню личных данных на компьютере Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a>Реализация

Ознакомьтесь с реализацией области задач на примере [надстройки Excel "Тенденции расходов банка WoodGrove" на JS](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) на сайте GitHub.

## <a name="see-also"></a>См. также

- [Fabric Core в надстройках Office](fabric-core.md)
- [Конструктивные шаблоны для надстроек Office](../design/ux-design-pattern-templates.md)
