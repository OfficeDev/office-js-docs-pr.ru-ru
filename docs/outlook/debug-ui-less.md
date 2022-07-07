---
title: Отладка надстройки Outlook без пользовательского интерфейса
description: Узнайте, как выполнить отладку надстройки Outlook без пользовательского интерфейса.
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: e46bdf15172f5224995b17c39df4ba60ca6380ad
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660209"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>Отладка надстройки Outlook без пользовательского интерфейса

В этой статье описывается, как использовать расширение отладчика надстроек Office в Visual Studio Code для отладки надстроек Outlook без пользовательского [интерфейса](add-in-commands-for-outlook.md#executing-a-javascript-function). Действия надстроек без пользовательского интерфейса инициируются с помощью кнопки команды надстройки на ленте. Дополнительные сведения о командах надстроек см. в командах [надстройки для Outlook](add-in-commands-for-outlook.md).

В этой статье предполагается, что у вас уже есть проект надстройки, который вы хотите отладить. Чтобы создать надстройку без пользовательского интерфейса для отладки, выполните действия, описанные в руководстве по созданию надстройки Outlook для создания [сообщения](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Пометка надстройки для отладки

Если вы использовали генератор [Yeoman](../develop/yeoman-generator-overview.md) для надстроек Office для создания проекта надстройки, перейдите к разделу [](#configure-and-run-the-debugger) "Настройка" и запустите раздел отладчика далее в этой статье. При запуске `npm start` `UseDirectDebugger` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` для сборки надстройки и запуска локального сервера команда также задает значение раздела реестра, чтобы пометить надстройку для отладки.

В противном случае, если для создания надстройки использовался другой инструмент, выполните следующие действия.

1. Перейдите к разделу `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` реестра. Замените `[Add-in ID]` его **\<Id\>** манифестом надстройки.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Задайте для ключа `UseDirectDebugger` значение `1`.

## <a name="configure-and-run-the-debugger"></a>Настройка и запуск отладчика

Теперь, когда вы включили отладку надстройки, вы можете настроить и запустить отладчик. Чтобы узнать, как это сделать, выберите один из следующих параметров, применимый к среде выполнения.

- Если надстройка работает в среде выполнения WebView, см. расширение отладчика надстроек [Microsoft Office для Visual Studio Code](../testing/debug-with-vs-extension.md).

- Если надстройка работает в среде выполнения Microsoft Edge Chromium WebView2, см. сведения об отладке надстроек [в Windows с помощью Visual Studio Code и Microsoft Edge WebView2 (](../testing/debug-desktop-using-edge-chromium.md)на Chromium).

## <a name="see-also"></a>См. также

- [Команды надстроек Outlook](add-in-commands-for-outlook.md)
- [Обзор отладки надстроек Office](../testing/debug-add-ins-overview.md)
- [Отладка надстройки Outlook на основе событий](debug-autolaunch.md)
