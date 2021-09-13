---
title: Диалоговые окна в надстройках Office
description: Узнайте о лучших практиках визуального оформления диалогов в Office надстройки.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 6e3dff8249e7d198712c0058f9876aa4806c7e08
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151062"
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
