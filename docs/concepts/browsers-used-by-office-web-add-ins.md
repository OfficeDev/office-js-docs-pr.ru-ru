---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 01/20/2021
localization_priority: Normal
ms.openlocfilehash: c540eece3b74bb043cc8f4921c7c774511b5a60a
ms.sourcegitcommit: 54d141cefb7bdc5f16330747d0ec8e8e2bd03e93
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/21/2021
ms.locfileid: "49916463"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, которые отображаются с помощью iFrame при работе в Office в Интернете и с помощью встроенных элементов управления браузера в Office для настольных и мобильных клиентов. Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и обдвижка поставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Работает ли надстройка в Office в Интернете, Microsoft 365 или Office 2013 без подписки или более поздней.

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|ОС|Версия Office|Edge WebView2 (на основе Chromium) установлен?|Браузер|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|любой|Office в Интернете|Неприменимо|Браузер, в котором открыт Office.|
|Mac|любой|Неприменимо|Safari|
|iOS|любой|Неприменимо|Safari|
|Android|любой|Неприменимо|Chrome|
|Windows 7, 8.1, 10 | Office 2013 или более поздней подписки|Всё равно|Internet Explorer 11|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _И_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| Всё равно|Microsoft Edge<sup>2, 3 с</sup> исходным WebView (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Нет |Microsoft Edge<sup>2, 3 с</sup> исходным WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Да<sup>4</sup>|  Microsoft Edge<sup>2, 3 с</sup> WebView2 (на основе Chromium) |

<sup>1.</sup> [Дополнительные](/officeupdates/update-history-office365-proplus-by-date) сведения см. на странице истории обновлений и о том, как найти версию клиента [Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) и канал обновления.

<sup>2</sup> Когда используется Microsoft Edge, экранный диктор Windows 10 (иногда называемый "устройством чтения с экрана") читает тег на странице, которая открывается в области `<title>` задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

<sup>3</sup> Если надстройка включает элемент в манифест, она использует Internet Explorer 11 независимо от версии Windows или `Runtimes` Microsoft 365. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

<sup>4.</sup> В дополнение к установке Microsoft Edge необходимо установить встраивляемый контроль WebView2, чтобы его можно было встраить в Office. Чтобы установить его, [см. microsoft Edge WebView2 / Встраить веб-содержимое ... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).


> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если у любого из пользователей вашей надстройки есть платформы, которые используют Internet Explorer 11, то для использования синтаксиса и функций ECMAScript 2015 или более поздней, у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называется ES6) или более поздней платформы JavaScript или TypeScript, а затем скомпилируете код в ES5 JavaScript с помощью компиляторов, таких как [esel](https://babeljs.io/) или [tsc.](https://www.typescriptlang.org/index.html)
> - Написание в ECMAScript 2015 или более [](https://wikipedia.org/wiki/Polyfill_(programming)) поздней платформе JavaScript, но также загрузка библиотеки полизаполнен, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать ваш код.
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение неполадок Microsoft Edge

### <a name="service-workers-are-not-working"></a>Сотрудники службы не работают

Надстройки Office не поддерживают службы, если используется [исходный веб-просмотр Microsoft Edge.](/microsoft-edge/hosting/webview) Они поддерживаются с помощью [Edge WebView2 на основе Chromium.](/microsoft-edge/hosting/webview2)

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) со значением `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Вывод ошибок при попытке скачать PDF-файл

Непосредственное скачивание BLOB-файлов в формате PDF в надстройке не поддерживается, если браузером является Edge. Обходным решением является создание простого веб-приложения, которое скачивает BLOB-файлы в формате PDF. В надстройки вызовите метод и `Office.context.ui.openBrowserWindow(url)` передайте URL-адрес веб-приложения. Веб-приложение откроется в окне браузера за пределами Office.

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
