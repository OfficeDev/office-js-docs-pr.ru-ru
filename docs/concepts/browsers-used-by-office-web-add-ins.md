---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: 6347bbe6b02befeabf0fcd6f04545f6d543871c8
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076184"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Office Надстройки — это веб-приложения, которые отображаются с помощью iFrames при работе в Office в Интернете и с помощью встроенных элементов управления браузером в Office для настольных и мобильных клиентов. Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и двигатель поставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Работает ли надстройка в Office в Интернете, Microsoft 365 или без подписки Office 2013 или более поздней.

> [!IMPORTANT]
> **Internet Explorer по-прежнему Office надстройки**
>
> Корпорация Майкрософт заканчивает поддержку Internet Explorer, но это не влияет на Office надстройки. Некоторые сочетания платформ и Office версий, включая все версии с одновкулярной покупкой до Office 2019 г., будут по-прежнему использовать управление веб-просмотром, которое поставляется с Internet Explorer 11, для пользования надстройки, как поводится в этой статье. Кроме того, поддержка этих комбинаций и, следовательно, internet Explorer по-прежнему требуется для надстройок, представленных [в AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Меняются *две* вещи:
>
> - AppSource больше не тестирует надстройки в Office в Интернете с помощью Internet Explorer в качестве браузера. Но AppSource по-прежнему тестирует комбинации  платформы и Office настольных версий, которые используют Internet Explorer.
> - Средство [Script Lab прекратит](../overview/explore-with-script-lab.md) работу в Internet Explorer в 2021 году.

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|OS|Версия Office|Edge WebView2 (Chromium на основе) установлен?|Браузер|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|любой|Office в Интернете|Неприменимо|Браузер, в котором открыт Office.|
|Mac|любой|Неприменимо|Safari|
|iOS|любой|Неприменимо|Safari|
|Android|любой|Неприменимо|Chrome|
|Windows 7, 8.1, 10 | подписка Office 2013 или более поздней|Всё равно|Internet Explorer 11|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _И_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| Всё равно|Microsoft Edge<sup>2, 3 с</sup> оригинальным WebView (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Нет |Microsoft Edge<sup>2, 3 с</sup> оригинальным WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Да<sup>4</sup>|  Microsoft Edge<sup>2</sup> с WebView2 (Chromium основе) |

<sup>1.</sup> [Дополнительные](/officeupdates/update-history-office365-proplus-by-date) сведения см. на странице история обновления и Office клиентской версии и канала обновления. [](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)

<sup>2</sup> Когда Microsoft Edge используется, Windows 10(иногда называемый "считыватель экрана") читает тег на странице, открываемой в области `<title>` задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

<sup>3</sup> Если надстройка включает элемент манифеста, она не будет использовать Microsoft Edge с исходным `<Runtimes>` WebView (EdgeHTML). Если условия использования Microsoft Edge WebView2 (Chromium на основе) выполнены, надстройка использует этот браузер. В противном случае он использует Internet Explorer 11 независимо от Windows или Microsoft 365 версии. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> Встраивляемый контроль WebView2 должен быть установлен в дополнение к установке Microsoft Edge, чтобы Office его встраить. Чтобы установить его, [см. Microsoft Edge WebView2 / Встраить веб-контент ... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если у любого из пользователей надстройки есть платформы, которые используют Internet Explorer 11, то для синтаксиса и функций ECMAScript 2015 или более позднего времени у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называемый ES6) или позже JavaScript, или в TypeScript, а затем скомпилировать код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Напишите в ECMAScript 2015 или более [](https://en.wikipedia.org/wiki/Polyfill_(programming)) поздний JavaScript, а также загрузите библиотеку полифильмов, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать код.
>
> Дополнительные сведения об этих параметрах см. в [меню Support Internet Explorer 11.](../develop/support-ie-11.md)
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение Microsoft Edge проблем

### <a name="service-workers-are-not-working"></a>Работники служб не работают

Office Надстройки не поддерживают сотрудников службы при Microsoft Edge WebView, [EdgeHTML.](https://en.wikipedia.org/wiki/EdgeHTML) Они поддерживаются с помощью [Chromium Edge WebView2.](/microsoft-edge/hosting/webview2)

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) со значением `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Получить ошибки при попытке скачивания PDF-файла

Непосредственное скачивание blobs в формате PDF-файлов в надстройке не поддерживается при браузере Edge. Обходным решением является создание простого веб-приложения, которое скачивает blobs в формате PDF-файлов. В надстройки позвоните `Office.context.ui.openBrowserWindow(url)` методу и передайте URL-адрес веб-приложения. Это откроет веб-приложение в окне браузера за пределами Office.

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
