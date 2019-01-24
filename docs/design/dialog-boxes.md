---
title: Диалоговые окна в надстройках Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 78a3419dd93f2a19e3addbeb5a77271b5b124680
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388404"
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

- [Ресурсы для разработки на сайте GitHub](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Объект Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog)


