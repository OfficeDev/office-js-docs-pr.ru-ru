---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 03/24/2021
localization_priority: Normal
ms.openlocfilehash: 489367231e1ed48e0bee6f0a32ccc47a8b39aed9
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604633"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, которые отображаются с помощью iFrames при работе в Office в Интернете и использовании встроенных элементов управления браузером в Office для настольных и мобильных клиентов. Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и двигатель поставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Работает ли надстройка в Office в Интернете, Microsoft 365 или Office 2013 или более поздней подписки.

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|OS|Версия Office|Edge WebView2 (на основе хрома) установлен?|Браузер|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|любой|Office в Интернете|Неприменимо|Браузер, в котором открыт Office.|
|Mac|любой|Неприменимо|Safari|
|iOS|любой|Неприменимо|Safari|
|Android|любой|Неприменимо|Chrome|
|Windows 7, 8.1, 10 | Office 2013 или более поздней подписки|Всё равно|Internet Explorer 11|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _И_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| Всё равно|Microsoft Edge<sup>2, 3 с</sup> оригинальным WebView (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Нет |Microsoft Edge<sup>2, 3 с</sup> оригинальным WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Да<sup>4</sup>|  Microsoft Edge<sup>2</sup> с WebView2 (на основе хрома) |

<sup>1.</sup> Дополнительные сведения см. на странице [история](/officeupdates/update-history-office365-proplus-by-date) обновления и поиске клиентской версии Office и [канала обновления.](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)

<sup>2</sup> Когда используется Microsoft Edge, рассказчик Windows 10 (иногда называемый "считывателем экрана") читает тег на странице, открываемой в области `<title>` задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

<sup>3</sup> Если ваша надстройка включает элемент манифеста, она не будет использовать Microsoft Edge с исходным `<Runtimes>` WebView (EdgeHTML). Если условия использования Microsoft Edge с WebView2 (на основе хрома) выполнены, надстройка использует этот браузер. В противном случае он использует Internet Explorer 11 независимо от версии Windows или Microsoft 365. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> Встраивляемый контроль WebView2 должен быть установлен в дополнение к установке Microsoft Edge, чтобы Office можно было встраить его. Чтобы установить его, см. [в веб-контенте Microsoft Edge WebView2 / Embed... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).




> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если у любого из пользователей надстройки есть платформы, которые используют Internet Explorer 11, то для синтаксиса и функций ECMAScript 2015 или более позднего времени у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называемый ES6) или позже JavaScript, или в TypeScript, а затем скомпилировать код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Напишите в ECMAScript 2015 или более [](https://en.wikipedia.org/wiki/Polyfill_(programming)) поздний JavaScript, а также загрузите библиотеку полифильмов, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать код.
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение неполадок с microsoft Edge

### <a name="service-workers-are-not-working"></a>Работники служб не работают

Надстройки Office не поддерживают сотрудников служб, когда используется оригинальный Microsoft Edge [WebView, EdgeHTML.](https://en.wikipedia.org/wiki/EdgeHTML) Они поддерживаются с [помощью edge WebView2 на основе хрома.](/microsoft-edge/hosting/webview2)

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
