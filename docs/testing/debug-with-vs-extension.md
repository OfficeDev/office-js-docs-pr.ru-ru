---
title: Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"
description: Используйте расширение Visual Studio кода Microsoft Office отладить надстройку Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237933"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"

Расширение Microsoft Office отладка надстройки для Visual Studio Code позволяет выполнить отладку надстройки Office в Microsoft Edge с помощью исходной времени работы WebView (EdgeHTML). Инструкции по отладке в Microsoft Edge WebView2 (на основе Chromium) см. [в этой статье](./debug-desktop-using-edge-chromium.md)

Этот режим отладки является динамическим, что позволяет устанавливать точки останова во время работы кода. Вы можете сразу увидеть изменения в коде, когда отладка подключена, без потери сеанса отладки. Изменения в коде также сохраняются, поэтому вы можете увидеть результаты нескольких изменений в коде. На следующем рисунке показано это расширение в действии.

![Расширение надстройки Office Addin Debugger Extension отладка раздела надстроек Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Необходимые компоненты

- [Visual Studio кода](https://code.visualstudio.com/) (должен запускаться от учетной записи администратора)
- [Node.js (версия 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

В этих инструкциях предполагается, что у вас есть опыт работы с командной строкой, вы понимаете базовый javaScript и создали проект надстройки Office перед использованием генератора Yo Office. Если вы еще не сделали этого, рассмотрите возможность посетить одно из наших учебников, например это руководство по [надстройки Excel Для Office.](../tutorials/excel-tutorial.md)

## <a name="install-and-use-the-debugger"></a>Установка и использование отладщика

1. Если вам нужно создать проект надстройки, создайте его с помощью генератора [Yo Office.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator) Следуйте подсказкам в командной строке, чтобы настроить проект. Вы можете выбрать любой язык или тип проекта в соответствии со своими потребностями.

> [!NOTE]
> Если у вас уже есть проект, пропустите шаг 1 и переходить к шагу 2.

2. Откройте командную подсказку от администратора.
   ![Параметры командной подсказки, в том числе "Запуск от администратора" в Windows 10](../images/run-as-administrator-vs-code.jpg)

3. Перейдите в каталог проекта.

4. Чтобы открыть проект в Visual Studio code от администратора, Visual Studio следующую команду.

```command&nbsp;line
code .
```

После Visual Studio кода перейдите в папку проекта вручную.

> [!TIP]
> Чтобы открыть Visual Studio Code от имени администратора, выберите параметр "Запуск от имени администратора" при открытии Visual Studio Code после его поиска в Windows. 

5. В VS Code выберите **CTRL + SHIFT + X,** чтобы открыть план "Расширения". Найщите расширение Microsoft Office надстройки и установите его.

6. В папке VSCODE проекта откройте файлlaunch.js **файла.** Добавьте в раздел следующий `configurations` код:

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

7. В разделе JSON, который вы только что скопировали, найдите раздел "URL". В этом URL-адресе необходимо заменить верхний регистр текста HOST на приложение, в которое размещена надстройка Office. Например, если ваша надстройка Office для Excel, url-адрес будет иметь значение https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0".

8. Откройте командную подсказку и убедитесь, что находитесь в корневой папке проекта. Запустите `npm start` команду, чтобы запустить сервер разработчиков. Когда надстройка загружается в клиенте Office, откройте области задач.

9. Вернись к Visual Studio Code и выберите "Просмотр > **Отлаки"** или введите **CTRL + SHIFT + D,** чтобы переключиться на представление отлаки.

10. В параметрах отлаки выберите **"Присоединение к надстройкам Office".** Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.

11. Установите точку останова в файле области задач проекта. Вы можете установить точки останова в VS Code, наведите курсор на строку кода и выберите красный круг.

![Красный круг отображается на строке кода в VS Code](../images/set-breakpoint.jpg)

12. Запустите надстройку. Вы увидите, что были сбиты точки останова, и можете проверить локальные переменные.

## <a name="see-also"></a>См. также

* [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)

* [Отладка надстроек с помощью средств разработчика в Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [Отладка надстроек в Windows с помощью Microsoft Edge WebView2 (на основе Chromium)](debug-desktop-using-edge-chromium.md)
