---
title: Тестирование Internet Explorer 11
description: Проверьте Office надстройки в Internet Explorer 11.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 994fde9ee255844f70f1d04f5898f7fffe0e3c5f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745318"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Проверьте Office надстройки в Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему Office надстройки**
>
> Корпорация Майкрософт заканчивает поддержку Internet Explorer, но это не Office надстройки. Некоторые комбинации платформ и Office версий, включая версии с одновековой покупкой до Office 2019 г., будут по-прежнему использовать управление веб-просмотром, которое поставляется с Internet Explorer 11 для пользования надстройки, как это объясняется в [браузерах](../concepts/browsers-used-by-office-web-add-ins.md), используемых Office надстройки. Кроме того, поддержка этих комбинаций и, следовательно, internet Explorer по-прежнему требуется для надстройок, представленных [в AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Меняются *две* вещи:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, AppSource больше не тестирует надстройки в Office в Интернете с помощью Internet Explorer в качестве браузера. Но AppSource по-прежнему тестирует комбинации платформы и Office настольных версий, которые используют Internet Explorer. 
> - Средство [Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.

Если вы планируете выставлять надстройку на рынок через AppSource или планируете поддерживать более старые версии Windows и Office, ваша надстройка должна работать в встраиваемом контроле браузера, основанном на Internet Explorer 11 (IE11). Вы можете использовать командную строку для перехода от более современных времен работы, используемых надстройки, к времени запуска Internet Explorer 11 для этого тестирования. Сведения о том, какие версии Windows и Office с помощью управления [веб-представлением](../concepts/browsers-used-by-office-web-add-ins.md) Internet Explorer 11, см. в Office надстройки.

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней части, у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называемый ES6) или позже JavaScript или в TypeScript, а затем составить код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Напишите в ECMAScript 2015 или более поздний JavaScript, а также [](https://en.wikipedia.org/wiki/Polyfill_(programming)) загрузите библиотеку полифильмов, например [core-js](https://github.com/zloirock/core-js), которая позволяет IE запускать код.
>
> Дополнительные сведения об этих параметрах см. в [меню Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение. Дополнительные дополнительные информации см. в добавлении Определение времени запуска надстройки в [Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> Office в Интернете не удается открыть в Internet Explorer 11, поэтому нельзя (и не нужно) тестировать надстройки на Office в Интернете internet Explorer.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Переключиться на веб-просмотр Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Существует два способа переключения веб-браузера Internet Explorer. Вы можете запустить простую команду в командной подсказке или установить версию Office, использующую Internet Explorer по умолчанию. Рекомендуем первый метод. Но второй вариант следует использовать в следующих сценариях.

- Ваш проект был разработан с Visual Studio и IIS. Это не node.js основе.
- Вы хотите быть абсолютно надежным в тестировании.
- Если по какой-либо причине средство командной строки не работает.

### <a name="switch-via-the-command-line"></a>Переключение через командную строку

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Установка версии Office, использующей Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>См. также

* [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)
* [Загрузка неопубликованных надстроек Office для тестирования](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Отладка надстроек с помощью средств разработчика для Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
* [Подключение отладчика из области задач](attach-debugger-from-task-pane.md)
