---
title: Тестирование Internet Explorer 11
description: Протестируйте надстройку Office в Internet Explorer 11.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9ab904a3b086990cb9b10e2f266ddacafb4cba94
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423330"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Тестирование надстройки Office в Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему используется в надстройки Office**
>
> В некоторых сочетаниях платформ и версий Office, включая версии Office 2019 с однофакторной покупкой, по-прежнему используется элемент управления WebView, который поставляется с Internet Explorer 11 для размещения надстроек, как описано в браузерах, используемых надстройки [Office](../concepts/browsers-used-by-office-web-add-ins.md). Рекомендуется (но не обязательно) продолжать поддерживать эти сочетания, по крайней мере минимально, предоставляя пользователям надстройки корректное сообщение об ошибке при запуске надстройки в веб-представлении Internet Explorer. Учитывайте следующие дополнительные моменты:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) больше не тестирует надстройки в Office в Интернете в качестве браузера.
> - AppSource по-прежнему тестирует сочетания версий платформы *и классических* версий Office, использующих Internet Explorer, однако предупреждение выдано только в том случае, если надстройка не поддерживает Internet Explorer. AppSource не отклоняет надстройку.
> - Средство [Script Lab больше](../overview/explore-with-script-lab.md) не поддерживает Internet Explorer.

Если вы планируете поддерживать более старые версии Windows и Office, надстройка должна работать в элементе управления встраиваемого браузера, основанном на Internet Explorer 11 (IE11). С помощью командной строки можно переключиться с более современных сред выполнения, используемых надстройки, на среду выполнения Internet Explorer 11 для этого тестирования. Сведения о том, в каких версиях Windows и Office используется элемент управления [веб-представлением](../concepts/browsers-used-by-office-web-add-ins.md) Internet Explorer 11, см. в разделе "Браузеры" надстроек Office.

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней версии, у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называемом ES6) или более поздней версии JavaScript или в TypeScript, а затем скомпилируете код в ES5 JavaScript с помощью компилятора, такого как [pythonel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Написание в ECMAScript 2015 или более поздней версии JavaScript, [](https://en.wikipedia.org/wiki/Polyfill_(programming)) но также загрузка библиотеки полизаполнения, например [core-js](https://github.com/zloirock/core-js), которая позволяет IE выполнять код.
>
> Дополнительные сведения об этих параметрах см. в [разделе "Поддержка Internet Explorer 11"](../develop/support-ie-11.md).
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение. Дополнительные сведения см. в статье "Определение во время выполнения", если надстройка [запущена в Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> - Office в Интернете невозможно открыть в Internet Explorer 11, поэтому вы не можете (и не нужно) тестировать надстройку в Office в Интернете Internet Explorer.
>
> - Для работы веб-надстроек Office необходимо отключить конфигурацию усиленной безопасности Internet Explorer (ESC). Если вы используете компьютер с Windows Server в качестве клиента при разработке надстроек, учитывайте, что конфигурация ESC включена по умолчанию в Windows Server.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Переключение на веб-представление Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Существует два способа переключения веб-представления Internet Explorer. Вы можете выполнить простую команду в командной строке или установить версию Office, использующую Internet Explorer по умолчанию. Мы рекомендуем использовать первый метод. Но второй вариант следует использовать в следующих сценариях.

- Ваш проект был разработан с помощью Visual Studio и IIS. Он не node.js основе.
- Вы хотите быть абсолютно надежными в тестировании.
- Если по какой-либо причине средство командной строки не работает.

### <a name="switch-via-the-command-line"></a>Переключение с помощью командной строки

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Установка версии Office, использующей Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>См. также

- [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)
- [Загрузка неопубликованных надстроек Office для тестирования](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Отладка надстроек с помощью средств разработчика для Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Подключение отладчика из области задач](attach-debugger-from-task-pane.md)
- [Среды выполнения в надстройки Office](runtimes.md)