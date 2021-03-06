---
title: Элементы пользовательского интерфейса Office для надстроек Office
description: Получите обзор различных элементов пользовательского интерфейса в Office надстройки.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5d0a1576d850f2291c28e6bb39554cbb0403f50b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076331"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Элементы пользовательского интерфейса Office для надстроек Office

Для расширения пользовательского интерфейса Office, в том числе команд надстроек и контейнеров HTML, можно использовать несколько типов элементов пользовательского интерфейса, которые полностью совместимы с Office и рядом платформ. Вы можете вставить пользовательский веб-код в любой из этих элементов.

На рисунке ниже приведены типы элементов пользовательского интерфейса Office, которые можно создать.

![Схема, показывающая команды надстройки в ленте, области задач и диалоговом окне/надстройке контента в Office документе.](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>Команды надстроек

Используйте [команды надстройки,](add-in-commands.md) чтобы добавить точки входа в надстройки к ленте Приложение Office. Команды запускают действия в надстройке путем выполнения кода JavaScript или запуска контейнера HTML. Можно создать два типа команд надстроек.

|Тип команды|Описание|
|:---------------|:--------------|
|Кнопки, меню и вкладки на ленте|Позволяют добавлять в Office пользовательские кнопки, меню (раскрывающиеся меню) или вкладки на ленте по умолчанию. Кнопки и меню используются для запуска действия в Office. Вкладки позволяют сгруппировать и упорядочить кнопки и меню.|
|Контекстные меню| Используются для расширения контекстного меню по умолчанию. Контекстные меню отображаются, когда пользователи щелкают правой кнопкой мыши текст в документе Office или таблице Excel.|

## <a name="html-containers"></a>Контейнеры HTML

Контейнеры HTML позволяют внедрить код пользовательского интерфейса на основе HTML в клиентах Office. Эти веб-страницы затем могут ссылаться на API JavaScript для Office для взаимодействия с содержимым в документе. Можно создать HTML-контейнеры трех типов.

|Контейнер HTML|Описание|
|:-----------------|:--------------|
|[Области задач](task-pane-add-ins.md)|Отображение собственного пользовательского интерфейса в правой части документа Office. Области задач позволяют пользователям взаимодействовать с вашей надстройкой, работая с документом Office.|
|[Контентные надстройки](content-add-ins.md)|Отображение пользовательского интерфейса, внедренного в документы Office. Контентные надстройки позволяют пользователям взаимодействовать с вашей надстройкой непосредственно в документе Office. Например, может понадобиться отобразить внешнее содержимое (видео или визуализации данных из других источников). |
|[Диалоговые окна](dialog-boxes.md)|Отображение пользовательского интерфейса в диалоговом окне, которое накладывается на документ Office. Используйте диалоговое окно для действий, которые требуют внимания и не требуют непосредственного взаимодействия с документом.|

## <a name="see-also"></a>См. также

- [Команды надстроек для Excel, Word и PowerPoint](add-in-commands.md)
- [Области задач](task-pane-add-ins.md)
- [Контентные надстройки](content-add-ins.md)
- [Диалоговые окна](dialog-boxes.md)
