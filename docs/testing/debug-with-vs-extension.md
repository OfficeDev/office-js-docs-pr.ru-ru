---
title: Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"
description: Чтобы отладить Office надстройки, Visual Studio Code надстройки Microsoft Office надстройки.
ms.date: 08/18/2021
localization_priority: Normal
ms.openlocfilehash: ba831cfabdefbf3829bb702bf21a70ddb499b972
ms.sourcegitcommit: 7ced26d588cca2231902bbba3f0032a0809e4a4a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2021
ms.locfileid: "58505672"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"

Расширение Microsoft Office надстройки для Visual Studio Code позволяет отладить Office надстройку Microsoft Edge с исходным временем работы webView (EdgeHTML). Инструкции по отладки Microsoft Edge WebView2 (Chromium основе) см. [в этой статье](./debug-desktop-using-edge-chromium.md)

Этот режим отладки динамический, что позволяет устанавливать точки разрыва во время работы кода. Вы можете видеть изменения в коде сразу же, когда отладка присоединена, все без потери сеанса отладки. Изменения кода также сохраняются, поэтому вы можете видеть результаты нескольких изменений в коде. На следующем изображении показано это расширение в действии.

![Office Расширение надстройки Debugger Extension, отладка раздела Excel надстроек.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Необходимые компоненты

- [Код Visual Studio](https://code.visualstudio.com/) (необходимо запускать от имени администратора)
- [Node.js (версия 10. или более поздняя)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

Эти инструкции предполагают, что вы имеете опыт использования командной строки, понимаете базовый JavaScript и создали проект Office надстройки перед использованием генератора Yo Office. Если вы еще не сделали этого раньше, рассмотрите возможность посещения одного из наших учебников, например Excel Office [надстройки](../tutorials/excel-tutorial.md).

## <a name="install-and-use-the-debugger"></a>Установка и использование отладчика

1. Если вам нужно создать проект надстройки, используйте генератор Yo Office для [создания этого проекта.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator) Следуйте подсказкам в командной строке, чтобы настроить проект. Вы можете выбрать любой язык или тип проекта в соответствии с вашими потребностями. В этом руководстве Excel надстройка области задач.

    > [!NOTE]
    > Если у вас уже есть проект, пропустить шаг 1 и перейти на шаг 2.

1. Откройте командную подсказку в качестве администратора.
   ![Параметры командной подсказки, в том числе "запуск в качестве администратора" в Windows 10.](../images/run-as-administrator-vs-code.jpg)

1. Перейдите к каталогу проектов.

1. Запустите следующую команду, чтобы открыть проект в Visual Studio Code в качестве администратора.

    ```command&nbsp;line
    code .
    ```

  После открытия Visual Studio Code перейдите вручную в папку проекта.

  > [!TIP]
  > Чтобы открыть Visual Studio Code администратора, выберите  запуск в качестве параметра администратора при открытии Visual Studio Code после поиска Windows.

1. Находясь в коде VS, нажмите клавиши **CTRL+SHIFT+X**, чтобы открыть меню расширений. Поиск расширения "Microsoft Office надстройки Debugger" и установка его.

1. В папке проекта . vscode проекта откройте файл **launch.json**. Добавьте в раздел следующий `configurations` код.

    ```JSON
    {
      "type": "office-addin",
      "request": "attach",
      "name": "Attach to Office Add-ins",
      "port": 9222,
      "trace": "verbose",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "webRoot": "${workspaceFolder}",
      "timeout": 45000
    }
    ```

1. В разделе JSON, который вы только что скопировали, найдите `"url"` свойство. В этом URL-адресе необходимо заменить текст **HOST** верхнего шкафа приложением, которое Office надстройки. Например, если Office надстройка для Excel, значение URL-адреса будет `"https://localhost:3000/taskpane.html?_host_Info=Excel$Win32$16.01$en-US$\$\$\$0"` .

1. Откройте командную подсказку и убедитесь, что вы находитесь в корневой папке проекта. Запустите `npm start` команду, чтобы запустить сервер разработчиков. Когда надстройка загружается в приложении Office, откройте области задач.

1. Вернись Visual Studio Code и выберите просмотр > отлаговка или введите **Ctrl+Shift+D,** чтобы перейти на отключаемую точку зрения. 

1. Из параметров отлаговки выберите **Attach to Office надстроек.** Выберите **F5** или **> запустить** отладку из меню, чтобы начать отладку.

1. Установите точку разлома в файле области задач проекта. Вы можете установить точки Visual Studio Code, зависая рядом с строкой кода и выбрав красный круг, который появится.

    ![Красный круг отображается на строке кода в Visual Studio Code.](../images/set-breakpoint.jpg)

1. Запустите надстройку. Вы увидите, что были поражены точки разрыва, и вы можете проверить локальные переменные.

## <a name="see-also"></a>См. также

- [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)

- [Отладка надстроек с помощью средств разработчика в Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [Отладка настроек в Windows с использованием Microsoft Edge WebView2 (на основе Chromium)](debug-desktop-using-edge-chromium.md)
