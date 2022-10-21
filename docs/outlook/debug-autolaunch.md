---
title: Отладка надстройки Outlook на основе событий
description: Узнайте, как выполнить отладку надстройки Outlook, которая реализует активацию на основе событий.
ms.topic: article
ms.date: 10/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: e8065c454bbe1587a6e5b7189a4522c229e9aed1
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664674"
---
# <a name="debug-your-event-based-outlook-add-in"></a>Отладка надстройки Outlook на основе событий

В этой статье приводятся рекомендации по отладке при реализации активации на основе [событий в](autolaunch.md) надстройке. Функция активации на основе событий была представлена в наборе обязательных элементов [1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), а дополнительные события теперь доступны в последующих наборах обязательных элементов. Дополнительные сведения см. в разделе ["Поддерживаемые события"](autolaunch.md#supported-events).

> [!IMPORTANT]
> Эта возможность отладки поддерживается только в Outlook для Windows с подпиской на Microsoft 365.

В этой статье рассматриваются основные этапы включения отладки.

- [Пометка надстройки для отладки](#mark-your-add-in-for-debugging)
- [Настройка Visual Studio Code](#configure-visual-studio-code)
- [Присоединение Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

Если вы использовали генератор Yeoman для надстроек Office для создания проекта надстройки (например, выполнив пошаговое руководство по активации на основе [событий),](autolaunch.md) следуйте параметру "Создать с генератором **Yeoman** " в этой статье. В противном случае выполните **другие** действия. Visual Studio Code должна быть версия не ниже 1.56.1.

## <a name="mark-your-add-in-for-debugging"></a>Пометка надстройки для отладки

1. Задайте раздел реестра `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`. Замените `[Add-in ID]` идентификатор надстройки из манифеста.

    - **XML-манифест**: используйте значение дочернего **\<Id\>** элемента корневого **\<OfficeApp\>** элемента.
    - **Манифест Teams (предварительная версия)**: используйте значение свойства id корневого анонимного `{ ... }` объекта.

    **Создан с помощью генератора Yeoman**. В окне командной строки перейдите к корню папки надстройки и выполните следующую команду.

    ```command&nbsp;line
    npm start
    ```

    Помимо создания кода и запуска локального сервера, `UseDirectDebugger` эта команда должна задать раздел реестра для этой надстройки `1`.

    **Другое**: добавьте раздел `UseDirectDebugger` реестра в раздел `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`. Замените `[Add-in ID]` его **\<Id\>** на манифест надстройки. Задайте для раздела реестра значение `1`.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Запустите Outlook или перезапустите его, если он уже открыт.
1. Создайте новое сообщение или встречу. Должно появиться диалоговое окно обработчика событий отладки. Пока *не* взаимодействуйте с диалогом.

    ![Диалоговое окно обработчика событий отладки в Windows.](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Настройка Visual Studio Code

### <a name="created-with-yeoman-generator"></a>Создано с помощью генератора Yeoman

1. Вернитесь в окно командной строки и откройте Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. В Visual Studio Code откройте файл **./.vscode/launch.json** и добавьте следующий фрагмент в список конфигураций. Сохраните изменения.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a>Другое

1. Создайте новую папку с **именем "Отладка** " (возможно, в **папке "Рабочий** стол").
1. Откройте Visual Studio Code.
1. Перейдите **в папку "** > **Открыть файл"**, перейдите к только что созданной папке, а затем выберите **"Выбрать папку"**.
1. На панели действий выберите "Выполнить **и отладка** " (CTRL+SHIFT+D).

    ![Значок запуска и отладки на панели действий.](../images/vs-code-debug.png)

1. Выберите **ссылку на файл launch.json** .

    ![Ссылка, расположенная в параметре запуска и отладки для создания файла launch.json в Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. В **раскрывающемся списке** "Выбор среды" выберите **"Edge: Launch** to create a launch.json file".
1. Добавьте следующий фрагмент в список конфигураций. Сохраните изменения.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a>Присоединение Visual Studio Code

1. Чтобы найтиbundle.jsнадстройки **, откройте** следующую папку в проводнике Windows **\<Id\>** и найдите ее (найти в манифесте).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Откройте папку с префиксом этого идентификатора и скопируйте полный путь. В Visual Studio Code **откройтеbundle.jsиз** этой папки. Шаблон пути к файлу должен выглядеть следующим образом:

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Поместите точки останова в bundle.js, где нужно остановить отладчик.
1. В **раскрывающемся списке DEBUG** **выберите прямую** отладку, а затем щелкните значок **"Начать отладку** ".

    ![Параметр прямой отладки, выбранный из параметров конфигурации в раскрывающемся Visual Studio Code отладки.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Отладка

1. Убедився, что отладчик подключен, вернитесь в Outlook и в диалоговом  окне обработчика на основе событий отладки нажмите кнопку "ОК **"**.

1. Теперь вы можете нажать точки останова в Visual Studio Code, что позволяет выполнять отладку кода активации на основе событий.

## <a name="stop-debugging"></a>Остановить отладку

Чтобы остановить отладку для остальной части текущего сеанса рабочего стола Outlook, в  диалоговом окне обработчика на основе событий отладки нажмите кнопку "**Отмена"**. Чтобы повторно включить отладку, перезапустите рабочий стол Outlook.

Чтобы диалоговое  окно обработчика на основе событий отладки не появлялись и не отлаживались для последующих сеансов Outlook, `0`удалите связанный раздел реестра или задайте для него значение : . `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`

## <a name="see-also"></a>См. также

- [Настройка надстройки Outlook для активации на основе событий](autolaunch.md)
- [Отладка надстройки с помощью журнала среды выполнения](../testing/runtime-logging.md#runtime-logging-on-windows)
