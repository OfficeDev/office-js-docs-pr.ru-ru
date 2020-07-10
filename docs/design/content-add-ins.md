---
title: Контентные надстройки Office
description: Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Excel или PowerPoint, что предоставляет пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f228ae8e7cca0426b0b43e31e38454029e4c7614
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093849"
---
# <a name="content-office-add-ins"></a>Контентные надстройки Office

Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Excel или PowerPoint. Контентные надстройки предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных. Используйте контентные надстройки, когда требуется внедрить функции непосредственно в документ.  

*Рисунок 1. Макет для контентных надстроек*

![Изображение, на котором показан типичный макет контентной надстройки.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Рекомендации

- Добавьте элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.
- Добавьте элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки (применимо только к надстройкам Excel и PowerPoint).

## <a name="variants"></a>Варианты

Размеры контентных надстроек для Excel и PowerPoint в Office для настольных ПК и Microsoft 365 указаны пользователем.

## <a name="personality-menu"></a>Меню личных данных

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

В Windows меню личных данных имеет размер 12 x 32 пикселей, как показано на изображении.

*Рисунок 2. Меню личных данных в Windows* 

![Изображение меню личных данных на компьютере с Windows](../images/personality-menu-win.png)


В Mac меню личных данных имеет размер 26 x 26 точек, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселей, как показано на изображении.

*Рисунок 3. Меню личных данных на Mac*

![Изображение меню личных данных на компьютере с Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Реализация

Пример реализации контентной надстройки для Excel: [Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) на сайте GitHub.

## <a name="support-considerations"></a>Что касается поддержки

- Проверьте, будет ли ваша надстройка Office работать на [конкретной платформе Office](../overview/office-add-in-availability.md).
- Чтобы надстройка могла читать и записывать данные в Excel или PowerPoint, может потребоваться добавление в список доверенных. Вы можете объявить нужный [уровень разрешений](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) для пользователя в манифесте надстройки.  
- Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.

## <a name="see-also"></a>См. также

- [Сведения о доступности элементов для надстроек Office, представленные с учетом ведущих приложений и платформ](../overview/office-add-in-availability.md)
- [Office UI Fabric в надстройках Office](../design/office-ui-fabric.md)
- [Конструктивные шаблоны для надстроек Office](../design/ux-design-pattern-templates.md)
- [Запрос разрешений на использование API в надстройках](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
