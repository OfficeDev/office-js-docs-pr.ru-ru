---
title: Диалоговые окна в надстройках Office
description: Узнайте о лучших практиках визуального оформления диалогов в Office надстройки.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 84af1bf7d5574ef87d66f801f7e1f7e74934601fcebcc9c273b9e9e40e8f4e38
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081967"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Диалоговые окна в надстройках Office

Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.

*Рисунок 1. Типичный макет диалогового окна*

![Типичная макетная схема диалогового окна, отображаемого в Office приложении.](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>Рекомендации

|Правильно|Неправильно|
|:-----|:--------|
|<ul><li>Укажите описательное название, содержащее имя надстройки и название текущей задачи.</li></ul>|<ul><li>Не включайте в него название вашей компании.</li></ul>|
||<ul><li>Не открывайте диалоговое окно, если этого не требует сценарий.</li></ul>|

## <a name="implementation"></a>Реализация

Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.

## <a name="see-also"></a>См. также

- [Объект Dialog](/javascript/api/office/office.dialog)
- [Конструктивные шаблоны для надстроек Office](../design/ux-design-pattern-templates.md)
