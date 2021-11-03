---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: e2660b51b0c25b15df5d957733631391401121ca
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681141"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Office Надстройки — это веб-приложения, которые отображаются с помощью iFrames при Office в Интернете. В Office для настольных и мобильных Office надстройки используют встроенный элемент управления браузером (также известный как веб-просмотр). Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и двигатель поставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Работает ли надстройка в Office в Интернете, Microsoft 365 или без подписки Office 2013 или более поздней.

> [!IMPORTANT]
> **Internet Explorer по-прежнему Office надстройки**
>
> Корпорация Майкрософт заканчивает поддержку Internet Explorer, но это не влияет на Office надстройки. Некоторые сочетания платформ и Office версий, включая версии с одновкулярной покупкой до Office 2019 г., будут по-прежнему использовать управление веб-просмотром, которое поставляется с Internet Explorer 11, для пользования надстройки, как поясняется в этой статье. Кроме того, поддержка этих комбинаций и, следовательно, internet Explorer по-прежнему требуется для надстройок, представленных [в AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Меняются *две* вещи:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, AppSource больше не тестирует надстройки в Office в Интернете с помощью Internet Explorer в качестве браузера. Но AppSource по-прежнему тестирует комбинации  платформы и Office настольных версий, которые используют Internet Explorer.
> - Средство [Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|OS|Версия Office|Edge WebView2 (Chromium на основе) установлен?|Браузер|
|:-----|:-----|:-----|:-----|
|любой|Office в Интернете|Неприменимо|Браузер, в котором открыт Office.<br>(Но обратите внимание, Office в Интернете не будет открываться в Internet Explorer.<br>Попытка сделать это открывает Office в Интернете edge.) |
|Mac|любой|Неприменимо|Safari|
|iOS|любой|Неприменимо|Safari|
|Android|любой|Неприменимо|Chrome.|
|Windows 7, 8.1, 10, 11 | подписка Office 2013 Office 2019 г.|Всё равно|Internet Explorer 11|
|Windows 10, 11 | подписка Office 2021 или более поздней|Да|Microsoft Edge<sup>1</sup> с WebView2 (Chromium основе)|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903,<br>Windows 11 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>2</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903,<br>Windows 11 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _И_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>2</sup>| Всё равно|Microsoft Edge<sup>1, 3 с</sup> оригинальным WebView (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903,<br>Окно 11 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>2</sup>| Нет |Microsoft Edge<sup>1, 3 с</sup> оригинальным WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>2</sup>| Да<sup>4</sup>|  Microsoft Edge<sup>1</sup> с WebView2 (Chromium основе) |

<sup>1</sup> При Microsoft Edge используется Windows(иногда называемый "считыватель экрана") считывателя тега на странице, открываемой в области `<title>` задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

<sup>2.</sup> [Дополнительные](/officeupdates/update-history-office365-proplus-by-date) сведения см. на странице история обновления и как найти Office клиентскую версию и канал обновления. [](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)

<sup>3</sup> Если надстройка включает элемент манифеста, она не будет использовать Microsoft Edge с исходным `<Runtimes>` WebView (EdgeHTML). Если условия использования Microsoft Edge WebView2 (Chromium на основе) выполнены, надстройка использует этот браузер. В противном случае он использует Internet Explorer 11 независимо от Windows или Microsoft 365 версии. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> В Windows до Windows 11 необходимо установить управление WebView2, чтобы Office его встраить. Он установлен с Microsoft 365 версии 2101 или более поздней версии, а также с Office 2021 или более поздней версии; но он не устанавливается автоматически с помощью Microsoft Edge. Если у вас есть более раная версия Microsoft 365 или разовая покупка Office, используйте инструкции по установке управления в [Microsoft Edge WebView2 / Embed веб-контента ... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). На Microsoft 365 сборки до 16.0.14326.xxxxx необходимо также создать ключ  реестраHKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2и установить его значение `dword:00000001` .

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если у любого из пользователей надстройки есть платформы, которые используют Internet Explorer 11, то для использования синтаксиса и функций ECMAScript 2015 или более поздней модели есть два варианта.
>
> - Напишите код в ECMAScript 2015 (также называемый ES6) или позже JavaScript, или в TypeScript, а затем скомпилировать код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Напишите в ECMAScript 2015 или более [](https://en.wikipedia.org/wiki/Polyfill_(programming)) поздний JavaScript, а также загрузите библиотеку полифильмов, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать код.
>
> Дополнительные сведения об этих параметрах см. в [меню Support Internet Explorer 11.](../develop/support-ie-11.md)
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение. Дополнительные дополнительные информации см. в добавлении Определить во время запуска, запущена ли надстройка [в Internet Explorer.](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение Microsoft Edge проблем

### <a name="service-workers-are-not-working"></a>Работники служб не работают

Office Надстройки не поддерживают сотрудников службы при Microsoft Edge WebView, [EdgeHTML.](https://en.wikipedia.org/wiki/EdgeHTML) Они поддерживаются с помощью [Chromium Edge WebView2.](/microsoft-edge/hosting/webview2)

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) со значением `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Получить ошибки при попытке скачивания PDF-файла

Непосредственное скачивание blobs в формате PDF-файлов в надстройке не поддерживается при браузере Edge. Обходным решением является создание простого веб-приложения, которое скачивает blobs в формате PDF-файлов. В надстройки позвоните `Office.context.ui.openBrowserWindow(url)` методу и передайте URL-адрес веб-приложения. Это откроет веб-приложение в окне браузера за пределами Office.

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
