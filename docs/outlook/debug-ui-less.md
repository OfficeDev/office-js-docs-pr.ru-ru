---
title: Отладка надстройки без Outlook пользовательского интерфейса
description: Узнайте, как выполнять отладку надстройки без пользовательского Outlook пользовательского интерфейса.
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 33aa36f86b7a163e650a23296b4c35aca7cb5492
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650713"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>Отладка надстройки без Outlook пользовательского интерфейса

В этой статье описывается, как использовать расширение Office отладчика надстроек в Visual Studio Code для отладки надстроек без пользовательского интерфейса [Outlook надстроек](add-in-commands-for-outlook.md#executing-a-javascript-function). Действия надстроек без пользовательского интерфейса инициируются с помощью кнопки команды надстройки на ленте. Дополнительные сведения о командах надстроек см. в [Outlook.](add-in-commands-for-outlook.md)

В этой статье предполагается, что у вас уже есть проект надстройки, который вы хотите отладить. Чтобы создать надстройку без пользовательского интерфейса для отладки, выполните действия, описанные в руководстве по созданию сообщения Outlook [надстройке](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Пометка надстройки для отладки

Если вы использовали генератор [Yeoman для Office](../develop/yeoman-generator-overview.md) надстроек для создания проекта надстройки, перейдите к разделу "Настройка и [](#configure-and-run-the-debugger) запуск отладчика" далее в этой статье. При запуске `npm start` `UseDirectDebugger` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` для сборки надстройки и запуска локального сервера команда также задает значение раздела реестра, чтобы пометить надстройку для отладки.

В противном случае, если для создания надстройки использовался другой инструмент, выполните следующие действия.

1. Перейдите к разделу `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` реестра. Замените `[Add-in ID]` **идентификатор** из манифеста надстройки.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Задайте для ключа `UseDirectDebugger` значение `1`.

## <a name="configure-and-run-the-debugger"></a>Настройка и запуск отладчика

Теперь, когда вы включили отладку надстройки, вы можете настроить и запустить отладчик. Чтобы узнать, как это сделать, выберите один из следующих параметров, применимый к среде выполнения.

- Если надстройка работает в среде выполнения WebView, см. Microsoft Office расширения отладчика надстройки для [Visual Studio Code.](../testing/debug-with-vs-extension.md)

- Если надстройка работает в среде выполнения Microsoft Edge Chromium WebView2, см. сведения об отладке надстроек в Windows с помощью [Visual Studio Code и Microsoft Edge WebView2 (](../testing/debug-desktop-using-edge-chromium.md)на основе Chromium).

## <a name="see-also"></a>См. также

- [Команды надстроек Outlook](add-in-commands-for-outlook.md)
- [Обзор отладки надстроек Office](../testing/debug-add-ins-overview.md)
- [Отладка надстройки Outlook на основе событий](debug-autolaunch.md)
