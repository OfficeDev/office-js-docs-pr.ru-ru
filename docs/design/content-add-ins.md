---
title: Контентные надстройки Office
description: Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Excel или PowerPoint, что предоставляет пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: f3dec371d1500d85125c8762bbc5e80f0cdfb571
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851287"
---
# <a name="content-office-add-ins"></a>Контентные надстройки Office

Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Excel или PowerPoint. Контентные надстройки предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных. Используйте контентные надстройки, когда требуется внедрить функции непосредственно в документ.  

*Рисунок 1. Макет для контентных надстроек*

![Изображение, на котором показан типичный макет контентной надстройки.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Рекомендации

- Добавьте элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.
- Добавьте элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки (применимо только к надстройкам Excel и PowerPoint).

## <a name="variants"></a>Варианты

Размеры контентных надстроек для Excel и PowerPoint в Office для настольных систем и Office 365 указывает пользователь.

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

- Проверьте, будет ли ваша надстройка Office работать на [конкретной платформе Office](/office/dev/add-ins/overview/office-add-in-availability). 
- Чтобы надстройка могла читать и записывать данные в Excel или PowerPoint, может потребоваться добавление в список доверенных. Вы можете объявить нужный [уровень разрешений](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) для пользователя в манифесте надстройки.  
- Контентные надстройки поддерживаются в Excel и PowerPoint в Office 2013 и более поздних версий. Если вы откроете надстройку в версии Office, которая не поддерживает веб-надстройки, вместо надстройки будет показано изображение.

## <a name="see-also"></a>См. также

- [Сведения о доступности элементов для надстроек Office, представленные с учетом ведущих приложений и платформ](/office/dev/add-ins/overview/office-add-in-availability)
- [Office UI Fabric в надстройках Office](/office/dev/add-ins/design/office-ui-fabric)
- [Конструктивные шаблоны для надстроек Office](/office/dev/add-ins/design/ux-design-pattern-templates)
- [Запрос разрешений на использование API в надстройках](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
