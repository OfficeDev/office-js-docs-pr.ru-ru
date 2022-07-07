---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1fedeb7f7e1e972a2a7fe4befa5a990ff8cc698d
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659656"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, которые отображаются с помощью iFrame при запуске в Office в Интернете. В Office для настольных и мобильных клиентов надстройки Office используют встроенный элемент управления браузером (также называемый веб-представлением). Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и обработчик предоставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Работает ли надстройка в Office в Интернете, Microsoft 365 или Office 2013 без подписки или более поздней версии.

> [!IMPORTANT]
> **Internet Explorer по-прежнему используется в надстройки Office**
>
> В некоторых сочетаниях платформ и версий Office, включая версии Office 2019 с однофакторной покупкой, по-прежнему используется элемент управления webview, который поставляется с Internet Explorer 11 для размещения надстроек, как описано в этой статье. Рекомендуется (но не обязательно) продолжать поддерживать эти сочетания, по крайней мере минимально, предоставляя пользователям надстройки корректное сообщение об ошибке при запуске надстройки в веб-представлении Internet Explorer. Учитывайте следующие дополнительные моменты:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) больше не тестирует надстройки в Office в Интернете в качестве браузера.
> - AppSource по-прежнему тестирует сочетания версий платформы *и классических* версий Office, использующих Internet Explorer, однако предупреждение выдано только в том случае, если надстройка не поддерживает Internet Explorer. AppSource не отклоняет надстройку.
> - Средство [Script Lab больше](../overview/explore-with-script-lab.md) не поддерживает Internet Explorer.
>
> Дополнительные сведения о поддержке Internet Explorer и настройке корректного сообщения об ошибке в надстройке см. в разделе ["Поддержка Internet Explorer 11"](../develop/support-ie-11.md).

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|ОС|Версия Office|Edge WebView2 (Chromium на основе) установлен?|Браузер|
|:-----|:-----|:-----|:-----|
|любой|Office в Интернете|Не применимо|Браузер, в котором открыт Office.<br>(Обратите внимание, Office в Интернете не будет открываться в Internet Explorer.<br>Попытка сделать это откроется Office в Интернете Edge.) |
|Mac|любой|Не применимо|Safari с WKWebView|
|iOS|любой|Не применимо|Safari с WKWebView|
|Android|любой|Не применимо|Chrome.|
|Windows 7, 8.1, 10, 11 | Office 2013 для Office 2019 без подписки|Всё равно|Internet Explorer 11|
|Windows 10, 11 | не из подписки Office 2021 или более поздней версии|Да|Microsoft Edge<sup>1</sup> с WebView2 (Chromium на основе)|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp; 16.0.11629<sup>2</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.11629&nbsp;_И_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| Всё равно|Microsoft Edge<sup>1, 3 с</sup> исходным WebView (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Окно 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Нет |Microsoft Edge<sup>1, 3 с</sup> исходным WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Да<sup>4</sup>|  Microsoft Edge<sup>1</sup> с WebView2 (Chromium на основе) |

<sup>1</sup> При использовании Microsoft Edge экранный диктор Windows (иногда называемый средством чтения с экрана) `<title>` считывает тег на странице, которая открывается в области задач. При использовании Internet Explorer 11 экранный диктор считывает строку заголовка области задач, **\<DisplayName\>** которая поступает из значения в манифесте надстройки.

<sup>2. Дополнительные</sup> [сведения см](/officeupdates/update-history-office365-proplus-by-date) . на странице журнала обновлений и о том, как найти версию клиента [Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) и канал обновления.

<sup>3</sup> . Если надстройка **\<Runtimes\>** содержит элемент в манифесте, она не будет использовать Microsoft Edge с исходным WebView (EdgeHTML). Если выполняются условия использования Microsoft Edge с WebView2 (Chromium на основе), надстройка использует этот браузер. В противном случае он использует Internet Explorer 11 независимо от версии Windows или Microsoft 365. Дополнительные сведения см. в статье [Runtimes](/javascript/api/manifest/runtimes).

<sup>4</sup> В версиях Windows, предшествующих Windows 11, необходимо установить элемент управления WebView2, чтобы Office его можно было внедрить. Он устанавливается с Microsoft 365 версии 2101 или более поздней версии и с однофакторной покупкой Office 2021 или более поздней версии, но не устанавливается автоматически в Microsoft Edge. Если у вас есть более раннюю версию Microsoft 365 или одновременная покупка Office, следуйте инструкциям по установке элемента управления в [Microsoft Edge WebView2 / Внедрение веб-содержимого... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). В сборках Microsoft 365 до версии 16.0.14326.xxxxx **** `dword:00000001`также необходимо создать раздел реестраHKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2и задать его значение.

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если у любого из пользователей надстройки есть платформы, использующие Internet Explorer 11, то для использования синтаксиса и функций ECMAScript 2015 или более поздней версии у вас есть два варианта.
>
> - Напишите код в ECMAScript 2015 (также называемом ES6) или более поздней версии JavaScript или в TypeScript, а затем скомпилируете код в ES5 JavaScript с помощью компилятора, такого как [pythonel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Написание в ECMAScript 2015 или более поздней версии JavaScript, [](https://en.wikipedia.org/wiki/Polyfill_(programming)) но также загрузка библиотеки полизаполнения, например [core-js](https://github.com/zloirock/core-js), которая позволяет IE выполнять код.
>
> Дополнительные сведения об этих параметрах см. в [разделе "Поддержка Internet Explorer 11"](../develop/support-ie-11.md).
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение. Дополнительные сведения см. в статье "Определение во время выполнения", если надстройка [запущена в Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение неполадок с Microsoft Edge

### <a name="service-workers-are-not-working"></a>Рабочие роли службы не работают

Надстройки Office не поддерживают service Workers при использовании исходного Microsoft Edge WebView, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML). Они поддерживаются в Chromium [Edge WebView2](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) со значением `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Получение ошибок при попытке скачать PDF-файл

Непосредственное скачивание больших двоичных объектов в виде PDF-файлов в надстройке не поддерживается, если браузером является Edge. Решением является создание простого веб-приложения, которое скачивает большие двоичные объекты в виде PDF-файлов. В надстройке вызовите метод `Office.context.ui.openBrowserWindow(url)` и передайте URL-адрес веб-приложения. Веб-приложение откроется в окне браузера за пределами Office.

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
