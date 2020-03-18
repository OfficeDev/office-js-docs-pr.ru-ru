---
title: Диалоговые окна в надстройках Office
description: Ознакомьтесь с рекомендациями по визуальному дизайну диалоговых окон в надстройках Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2f3b25fac7f12494e6b5a1e0a32e72baa345e978
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717196"
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

- [Объект Dialog](/javascript/api/office/office.dialog)
- [Конструктивные шаблоны для надстроек Office](../design/ux-design-pattern-templates.md)
