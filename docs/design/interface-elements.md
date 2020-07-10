---
title: Элементы пользовательского интерфейса Office для надстроек Office
description: Получите обзор различных типов элементов пользовательского интерфейса в надстройке Office.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5b9907924c674ed9db2294621123c394419d0c12
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093765"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Элементы пользовательского интерфейса Office для надстроек Office

You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.

На рисунке ниже приведены типы элементов пользовательского интерфейса Office, которые можно создать.

![Изображение с командами надстроек на ленте, областью задач и диалоговым окном в документе Office](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>Команды надстроек

Используйте [команды надстройки](add-in-commands.md) для добавления точек входа в надстройку на ленту приложения Office. Команды запускают действия в надстройке путем выполнения кода JavaScript или запуска контейнера HTML. Можно создать два типа команд надстроек.

|**Тип команды**|**Описание**|
|:---------------|:--------------|
|Кнопки, меню и вкладки на ленте|Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.|
|Контекстные меню| Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.| 

## <a name="html-containers"></a>Контейнеры HTML

Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.

|**Контейнер HTML**|**Описание**|
|:-----------------|:--------------|
|[Области задач](task-pane-add-ins.md)|Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.|
|[Контентные надстройки](content-add-ins.md)|Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources. |
|[Диалоговые окна](dialog-boxes.md)|Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.|

## <a name="see-also"></a>См. также

- [Команды надстроек для Excel, Word и PowerPoint](add-in-commands.md)
- [Области задач](task-pane-add-ins.md)
- [Контентные надстройки](content-add-ins.md)
- [Диалоговые окна](dialog-boxes.md)
