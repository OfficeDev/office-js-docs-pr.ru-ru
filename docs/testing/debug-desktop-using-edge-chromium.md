---
title: Отладка настроек в Windows с использованием Microsoft Edge WebView2 (на основе Chromium)
description: Узнайте, как осуществлять отладку надстроек Office, в которых используется Microsoft Edge WebView2 (на основе Chromium) с помощью отладчика для расширения Microsoft Edge в коде VS.
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 6a62718147fbb5d2e8a6819066425737d853cbf0
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350178"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>Отладка надстроек в Windows с помощью Edge Chromium WebView2

Надстройки Office, работающие в Windows, могут использовать отладчик для расширения Microsoft Edge в коде VS для отладки среды Edge Chromium WebView2.

## <a name="prerequisites"></a>Необходимые компоненты

- [Код Visual Studio](https://code.visualstudio.com/) (необходимо запускать от имени администратора)
- [Node.js (версия 10. или более поздняя)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge Chromium доступна участникам программы предварительной оценки Windows](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a>Установка и использование отладчика

1. Создайте проект с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Для этого можно использовать любые краткие руководства по началу работы, например [Краткое руководство по надстройкам Outlook](../quickstarts/outlook-quickstart.md).

    > [!TIP]
    > Если вы не используете надстройку, основанную на генераторе Yeoman, необходимо настроить ключ реестра. В корневой папке проекта выполните указанные ниже действия в командной строке: `office-add-in-debugging start <your manifest path>`.

1. Откройте проект в VS Code. Находясь в коде VS, нажмите **CTRL + SHIFT + X**, чтобы открыть меню расширений. Выполните поиск расширения "Debugger для Microsoft Edge" и установите его.

1. В папке проекта **. vscode** проекта откройте файл **launch.json**. Добавьте указанный ниже код в раздел конфигураций.

      ```JSON
        {
          "name": "Debug Office Add-in (Edge Chromium)",
          "type": "edge",
          "request": "attach",
          "useWebView": "advanced",
          "port": 9229,
          "timeout": 600000,
          "webRoot": "${workspaceRoot}",
        },
      ```

1. Чтобы перейти к представлению отладки, нажмите **Просмотр> Отладка** или введите **CTRL + SHIFT + D**.

1. В разделе параметров отладки выберите параметр Edge Chromium для ведущего приложения, например **классического приложения Excel (Edge Chromium)**. Чтобы начать отладку, нажмите **F5** или выберите **Отладка > Начать отладку** в меню.

1. Теперь надстройка готова к использованию в ведущем приложении, таком как Excel. Нажмите кнопку **Показать область задач** или выполнить другие дополнительные команды надстройки. Появится диалоговое окно подтверждения действия с надписью

    > WebView Stop On Load.
    > Чтобы выполнить отладку WebView, вложите код VS в экземпляр WebView с помощью отладчика Microsoft для Edge и нажмите кнопку ОК. Чтобы предотвратить появление диалогового окна в дальнейшем, нажмите кнопку"Отмена".

    Нажмите **ОК**.

    > [!NOTE]
    > После нажатия кнопки **Отмена** диалоговое окно не будет отображаться в процессе работы с этим экземпляром надстройки. Однако при перезапуске надстройки диалоговое окно снова появится.

1. Теперь можно задать точки останова в коде проекта и выполнить отладку.

## <a name="see-also"></a>См. также

- [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)
- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)
- [Подключение отладчика из области задач](attach-debugger-from-task-pane.md)