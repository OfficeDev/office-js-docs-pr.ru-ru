---
title: Команды функций отладки в надстройки Outlook
description: Узнайте, как выполнять отладку команд функций в надстройки Outlook.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6189824fd526d48321b355c9b306fa5ef732f411
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797591"
---
# <a name="debug-function-commands-in-outlook-add-ins"></a>Команды функций отладки в надстройки Outlook

> [!NOTE]
> Метод, описанный в этой статье, можно использовать только на компьютере разработки Windows. Если вы разрабатываете на компьютере Mac, см. команды [функции отладки](../testing/debug-function-command.md).

В этой статье описывается, как использовать расширение отладчика надстройки Office в Visual Studio Code для отладки команд [функций](add-in-commands-for-outlook.md#run-a-function-command). Команды функций инициируются с помощью кнопки команды надстройки на ленте. Дополнительные сведения о командах надстроек см. в командах [надстройки для Outlook](add-in-commands-for-outlook.md).

В этой статье предполагается, что у вас уже есть проект надстройки, который вы хотите отладить. Чтобы создать надстройку с командой функции для отладки, выполните действия, описанные в руководстве по созданию надстройки Outlook для создания [сообщения](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Пометка надстройки для отладки

Если вы использовали генератор [Yeoman](../develop/yeoman-generator-overview.md) для надстроек Office для создания проекта надстройки, перейдите к разделу [](#configure-and-run-the-debugger) "Настройка" и запустите раздел отладчика далее в этой статье. При запуске `npm start` `UseDirectDebugger` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` для сборки надстройки и запуска локального сервера команда также задает значение раздела реестра, чтобы пометить надстройку для отладки.

В противном случае, если для создания надстройки использовался другой инструмент, выполните следующие действия.

1. Перейдите к разделу `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` реестра. Замените `[Add-in ID]` его **\<Id\>** манифестом надстройки.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Задайте для ключа `UseDirectDebugger` значение `1`.

## <a name="configure-and-run-the-debugger"></a>Настройка и запуск отладчика

Теперь, когда вы включили отладку надстройки, вы можете настроить и запустить отладчик. Чтобы узнать, как это сделать, выберите один из следующих параметров, применимый к элементу управления webview. Сведения о том, как определить, какой элемент управления [веб-представление](../concepts/browsers-used-by-office-web-add-ins.md) используется на компьютере разработки, см. в разделе "Браузеры" надстроек Office.

- Если надстройка работает во внедренном элементе управления webview из устаревшей версии Edge (EdgeHTML), см. расширение отладчика надстройки [Microsoft Office для Visual Studio Code](../testing/debug-with-vs-extension.md).

- Если надстройка выполняется во внедренном элементе управления webview из Microsoft Edge Chromium (WebView2), см. раздел "Отладка надстроек в Windows с помощью [Visual Studio Code и Microsoft Edge WebView2 (на Chromium)"](../testing/debug-desktop-using-edge-chromium.md).

## <a name="see-also"></a>См. также

- [Команды надстроек Outlook](add-in-commands-for-outlook.md)
- [Обзор отладки надстроек Office](../testing/debug-add-ins-overview.md)
- [Отладка надстройки Outlook на основе событий](debug-autolaunch.md)
