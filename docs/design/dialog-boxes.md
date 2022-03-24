---
title: Диалоговые окна в надстройках Office
description: Узнайте о лучших практиках визуального оформления диалогов в Office надстройки.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 623a94b5cd0fdd398de2e61cb779f5d256c46d76
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743254"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Диалоговые окна в надстройках Office

Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.

*Рисунок 1. Типичный макет диалогового окна*

![Типичная макетная схема диалоговое окно, отображаемого в Office приложении.](../images/overview-with-app-dialog.png)

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
