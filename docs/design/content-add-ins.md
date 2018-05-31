---
title: Контентные надстройки Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd0dcea7a3f37175a48946fc9dcd61d2b89f9c08
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437264"
---
# <a name="content-office-add-ins"></a>Контентные надстройки Office

Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Word, Excel и PowerPoint. Контентные надстройки предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных. Используйте контентные надстройки, когда нужно внедрить функции прямо в документ.  

*Рисунок 1. Макет для контентных надстроек*

![Изображение, на котором показан типичный макет контентной надстройки.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Рекомендации

- Включите элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.
- Включите элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки (применимо только к надстройкам Word, Excel и PowerPoint).

## <a name="variants"></a>Варианты

Размеры контентных надстроек для Word, Excel и PowerPoint в Office 2016 для настольных систем и Office 365 указывает пользователь.

## <a name="personality-menu"></a>Меню личных данных

Меню личных данных могут перекрывать элементы навигации и управления, расположенные в правой верхней части надстройки. Ниже указаны текущие размеры меню личных данных в Windows и Mac.

В Windows меню личных данных имеет размер 12 x 32 пикселей, как показано на изображении.

*Рисунок 2. Меню личных данных в Windows* 

![Изображение меню личных данных на компьютере с Windows](../images/personality-menu-win.png)


В Mac меню личных данных имеет размер 26 x 26 точек, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселей, как показано на изображении.

*Рисунок 3. Меню личных данных на Mac*

![Изображение меню личных данных на компьютере с Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Реализация

Пример реализации контентной надстройки для Excel: [Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) на сайте GitHub.

## <a name="support-considerations"></a>Что касается поддержки
- Проверьте, будет ли ваша надстройка Office работать на [конкретной платформе Office](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability). 
- Чтобы надстройка могла читать и записывать данные в Excel или PowerPoint, может потребоваться добавление в список доверия сертификатов. Вы можете объявить нужный [уровень разрешений](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) в манифесте надстройки.  
- Надстройки контента поддерживаются в Excel и PowerPoint в Office 2013 и более поздних версиях. Если вы откроете надстройку в версии Office, которая не поддерживает веб-надстройки, вместо надстройки будет показано изображение.

## <a name="see-also"></a>См. также
- [Сведения о доступности элементов для надстроек Office, представленные с учетом ведущих приложений и платформ](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)
- [Office UI Fabric в надстройках Office](https://docs.microsoft.com/en-us/office/dev/add-ins/design/office-ui-fabric) 
- [Шаблоны проектирования взаимодействия для надстроек Office](https://docs.microsoft.com/en-us/office/dev/add-ins/design/ux-design-patterns)
- [Запрашивание разрешений на использование API в контентных надстройках и надстройках области задач](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
