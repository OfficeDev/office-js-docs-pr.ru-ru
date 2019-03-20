---
title: Диалоговые окна в надстройках Office
description: ''
ms.date: 02/28/2019
localization_priority: Priority
ms.openlocfilehash: 3638006c30515a1fcae93ccfdbd43e0e92005c37
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691106"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Диалоговые окна в надстройках Office
 
Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.

*Рисунок 1. Типичный макет диалогового окна*

![Изображение, на котором показан типичный макет диалогового окна](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>Рекомендации

|**Рекомендуется**|**Не рекомендуется**|
|:-----|:--------|
|<ul><li>Укажите описательное название, содержащее имя надстройки и название текущей задачи.</li></ul>|<ul><li>Не включайте в него название вашей компании.</li></ul>|
||<ul><li>Не открывайте диалоговое окно, если этого не требует сценарий.</li></ul>|

## <a name="implementation"></a>Реализация

Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.

## <a name="see-also"></a>См. также

- [Объект Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [Конструктивные шаблоны для надстроек Office](../design/ux-design-pattern-templates.md)


