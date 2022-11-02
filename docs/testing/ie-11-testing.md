---
title: Тестирование Internet Explorer 11
description: Протестируйте надстройку Office в Internet Explorer 11.
ms.date: 10/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: f5e962bb615849b4944be2bee3f14006b0c9289e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810366"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Тестирование надстройки Office в Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему используется в надстройках Office**
>
> Некоторые сочетания платформ и версий Office, включая бессрочные версии Office 2019, по-прежнему используют элемент управления webview, который поставляется с Internet Explorer 11, для размещения надстроек, как описано в [статье Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md). Мы рекомендуем (но не требовать), чтобы вы продолжали поддерживать эти сочетания, по крайней мере в минимальном виде, предоставляя пользователям надстройки корректное сообщение о сбое при запуске надстройки в веб-представлении Internet Explorer. Помните о следующих дополнительных моментах:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) больше не тестирует надстройки в Office в Интернете использует Internet Explorer в качестве браузера.
> - AppSource по-прежнему тестирует сочетание версий платформы и *классических* версий Office, использующих Internet Explorer, однако выдает предупреждение только в том случае, если надстройка не поддерживает Internet Explorer. Надстройка не отклоняется AppSource.
> - [Средство Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.

Если вы планируете поддерживать более старые версии Windows и Office, надстройка должна работать во встраиваемом элементе управления браузера, основанном на Internet Explorer 11 (IE11). Командную строку можно использовать для переключения с более современных сред выполнения, используемых надстройками, на среду выполнения Internet Explorer 11 для этого тестирования. Сведения о том, какие версии Windows и Office используют элемент управления веб-представлением Internet Explorer 11, см. [в статье Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней версии, у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называемом ES6) или более поздней версии JavaScript или в TypeScript, а затем скомпилируйте код в ES5 JavaScript с помощью компилятора, например [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Напишите в ECMAScript 2015 или более поздней версии JavaScript, но также загрузите библиотеку [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) , например [core-js](https://github.com/zloirock/core-js) , которая позволяет IE выполнять код.
>
> Дополнительные сведения об этих параметрах см. в разделе [Поддержка Internet Explorer 11](../develop/support-ie-11.md).
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение. Дополнительные сведения см [. в статье Определение надстройки во время выполнения в Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> - Office в Интернете нельзя открыть в Internet Explorer 11, поэтому вы не можете (и не нужно) тестировать надстройку на Office в Интернете с помощью Internet Explorer.
>
> - Для работы веб-надстроек Office необходимо отключить конфигурацию усиленной безопасности Internet Explorer (ESC). Если вы используете компьютер с Windows Server в качестве клиента при разработке надстроек, учитывайте, что конфигурация ESC включена по умолчанию в Windows Server.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Переключение на веб-представление Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Переключить веб-представление Internet Explorer можно двумя способами. Можно выполнить простую команду в командной строке или установить версию Office, которая по умолчанию использует Internet Explorer. Мы рекомендуем использовать первый метод. Но второй следует использовать в следующих сценариях.

- Проект был разработан с помощью Visual Studio и IIS. Это не node.js.
- Вы хотите быть абсолютно надежным в тестировании.
- Вы не можете использовать канал бета-версии для Microsoft 365 на компьютере разработки.
- Вы разрабатываете на Компьютере Mac. 
- Если по какой-либо причине программа командной строки не работает.

### <a name="switch-via-the-command-line"></a>Переключение с помощью командной строки

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Установка версии Office, использующего Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>См. также

- [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)
- [Загрузка неопубликованных надстроек Office для тестирования](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Отладка надстроек с помощью средств разработчика для Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Подключение отладчика из области задач](attach-debugger-from-task-pane.md)
- [Среды выполнения в надстройках Office](runtimes.md)