---
title: Отламывка надстройки Outlook событий (предварительный просмотр)
description: Узнайте, как отлагировать Outlook надстройки, которая реализует активацию на основе событий.
ms.topic: article
ms.date: 05/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: ebf469ec15948ae2daf693bc7fda692367d70bec
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154176"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a>Отламывка надстройки Outlook событий (предварительный просмотр)

В этой статье содержится руководство по отладки при реализации активации на основе событий [в](autolaunch.md) надстройки. Функция активации на основе событий в настоящее время находится в предварительном режиме.

> [!IMPORTANT]
> Эта возможность отладки поддерживается только для предварительного просмотра Outlook в Windows с Microsoft 365 подпиской. Дополнительные сведения см. в разделе [Отладка](#preview-debugging-for-the-event-based-activation-feature) предварительного просмотра для раздела функции активации на основе событий в этой статье.

В этой статье мы обсудим основные этапы, позволяющие отладку.

- [Пометить надстройку для отладки](#mark-your-add-in-for-debugging)
- [Настройка Visual Studio Code](#configure-visual-studio-code)
- [Прикрепить Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

У вас есть несколько вариантов создания проекта надстройки. В зависимости от используемого варианта действия могут отличаться. Если вы использовали генератор Yeoman для Office надстроек для создания проекта надстройки (например, с помощью погона активации на основе [событий),](autolaunch.md)выполните  действия **yo office,** в противном случае выполните другие действия. Visual Studio Code должна быть по крайней мере версия 1.56.1.

## <a name="preview-debugging-for-the-event-based-activation-feature"></a>Предварительная отладка функции активации на основе событий

Мы приглашаем вас попробовать возможности отладки для функции активации на основе событий! Дайте нам знать о ваших сценариях и о том, как мы можем улучшить ситуацию, GitHub с помощью GitHub (см. раздел **Обратная** связь в конце этой страницы).

Чтобы просмотреть эту возможность Outlook на Windows, минимальная требуемая сборка составляет 16.0.13729.20000. Чтобы получить доступ Office бета-версий, присоединитесь к [программе Office Insider.](https://insider.office.com)

## <a name="mark-your-add-in-for-debugging"></a>Пометить надстройку для отладки

1. Установите ключ `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` реестра. `[Add-in ID]` является **Id** в манифесте надстройки.

    **Yo office.** В окне командной строки перейдите к корневой папке надстройки и запустите следующую команду.

    ```command&nbsp;line
    npm start
    ```

    В дополнение к построению кода и запуску локального сервера эта команда должна установить ключ реестра для этой `UseDirectDebugger` надстройки. `1`

    **Другие:** Добавьте `UseDirectDebugger` ключ реестра под `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` . `[Add-in ID]`Замените **id из** манифеста надстройки. Установите ключ `1` реестра.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Запустите Outlook (или перезапустите Outlook, если он уже открыт).
1. Составить новое сообщение или назначение. Вы должны увидеть следующий диалог. Пока *не* взаимодействуйте с диалогом.

    ![Снимок экрана диалогового обработера событий на основе отладки.](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Настройка Visual Studio Code

### <a name="yo-office"></a>yo office

1. В окне командной строки откройте Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. В Visual Studio Code откройте **файл ./.vscode/launch.json** и добавьте следующий отрывок в список конфигураций. Сохраните изменения.

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

### <a name="other"></a>Прочее

1. Создайте новую папку под названием **Отладка** (возможно, в **папке Desktop).**
1. Откройте Visual Studio Code.
1. Перейдите **к**  >  **открытой папке File Open,** перейдите к только что созданной папке, а затем выберите **Выберите папку**.
1. В панели Действия выберите элемент **Отлаговка** (Ctrl+Shift+D).

    ![Снимок экрана значка Отлаговка в панели действий.](../images/vs-code-debug.png)

1. Выберите **ссылку на файл launch.json.**

    ![Снимок экрана ссылки для создания файла launch.json в Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. В **отсеве Выберите среду** выберите **Edge: Запуск** для создания файла launch.json.
1. Добавьте следующий отрывок в список конфигураций. Сохраните изменения.

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

## <a name="attach-visual-studio-code"></a>Прикрепить Visual Studio Code

1. Чтобы найтиbundle.jsнадстройки, **** откройте следующую папку в Windows Explorer и найдите **id** надстройки (найден в манифесте).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Откройте префикс папки с этим ID и скопируйте ее полный путь. В Visual Studio Code откройте **bundle.js** из этой папки. Шаблон пути файла должен быть следующим:

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Размыть точки bundle.js, где нужно остановить отладка.
1. В **отсеве DEBUG** выберите имя **Direct Debugging**, а затем выберите **Выполнить**.

    ![Снимок экрана выбора прямого отладки из параметров конфигурации в Visual Studio Code отладки.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Отладка

1. После подтверждения того, что отладка присоединена, вернись в  Outlook и в диалоговом окне обработник на основе событий отладки выберите **ОК** .

1. Теперь вы можете поразить точки Visual Studio Code, что позволит отключить код активации на основе событий.

## <a name="stop-debugging"></a>Остановка отладки

Чтобы остановить отладку для остальной части текущего сеанса  Outlook рабочего стола, в диалоговом оклините обработите для отладки событий выберите **Отмена**. Чтобы повторно включить отладку, перезапустите Outlook рабочего стола.

Чтобы предотвратить  отладку диалогового обработика событий на основе отладки и остановить отладку для последующих сеансов Outlook, удалите связанный ключ реестра или установите его `0` значение: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .

## <a name="see-also"></a>Дополнительные материалы

- [Настройка надстройки Outlook для активации на основе событий](autolaunch.md)
- [Отладка надстройки с помощью журнала среды выполнения](../testing/runtime-logging.md#runtime-logging-on-windows)
