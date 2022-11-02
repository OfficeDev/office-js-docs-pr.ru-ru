---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 09/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: a75cab613605760e774f8b2a163172e4ec6cb5bd
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810157"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, которые отображаются с помощью iFrames при запуске в Office в Интернете. В Office для классических и мобильных клиентов надстройки Office используют встроенный элемент управления браузера (также известный как веб-представление). Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и обработчик предоставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Выполняется ли надстройка в Office в Интернете, в Office, скачанном из подписки Microsoft 365, или в Office 2013 с бессрочным сроком действия или более поздней версии.
- В бессрочных версиях Office в Windows, независимо от того, работает ли надстройка в варианте "розничная" или "корпоративная лицензия".

> [!NOTE]
> В этой статье предполагается, что надстройка выполняется в документе, который *не* защищен [с помощью windows Information Protection (WIP).](/windows/uwp/enterprise/wip-hub) Для документов, защищенных WIP, существуют некоторые исключения из сведений, приведенных в этой статье. Дополнительные сведения см. в статье Документы, [защищенные WIP](#wip-protected-documents).

> [!IMPORTANT]
> **Internet Explorer по-прежнему используется в надстройках Office**
>
> Некоторые сочетания платформ и версий Office, включая корпоративные бессрочные версии до Office 2019, по-прежнему используют элемент управления webview, который поставляется с Internet Explorer 11, для размещения надстроек, как описано в этой статье. Мы рекомендуем (но не требовать), чтобы вы продолжали поддерживать эти сочетания, по крайней мере в минимальном виде, предоставляя пользователям надстройки корректное сообщение о сбое при запуске надстройки в веб-представлении Internet Explorer. Помните о следующих дополнительных моментах:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) больше не тестирует надстройки в Office в Интернете использует Internet Explorer в качестве браузера.
> - AppSource по-прежнему тестирует сочетание версий платформы и *классических* версий Office, использующих Internet Explorer. Однако он выдает предупреждение только в том случае, если надстройка не поддерживает Internet Explorer. Надстройка не отклоняется AppSource.
> - [Средство Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.
>
> Дополнительные сведения о поддержке Internet Explorer и настройке корректного сообщения об ошибке в надстройке см. в разделе [Поддержка Internet Explorer 11](../develop/support-ie-11.md).

В следующих разделах указывается, какой браузер используется для различных платформ и операционных систем.

## <a name="non-windows-platforms"></a>Платформы, отличные от Windows

Для этих платформ только платформа определяет используемый браузер.

|ОС|Версия Office|Браузер|
|:-----|:-----|:-----|
|любой|Office в Интернете|Браузер, в котором открыт Office.<br>(Но обратите внимание, что Office в Интернете не открывается в Internet Explorer.<br>При попытке сделать это откроется Office в Интернете в Edge.) |
|Mac|любой|Safari с WKWebView|
|iOS|любой|Safari с WKWebView|
|Android|любой|Chrome|

## <a name="perpetual-versions-of-office-on-windows"></a>Бессрочные версии Office в Windows

Для бессрочных версий Office в Windows используемый браузер определяется версией Office, независимо от того, является ли лицензия розничной или корпоративной, а также установлена ли edge WebView2 (на основе Chromium). Версия Windows не имеет значения, но обратите внимание, что веб-надстройки Office не поддерживаются в версиях, предшествующих Windows 7, и Office 2021 не поддерживается в версиях, предшествующих Windows 10.

Чтобы определить, является ли Office 2016 или Office 2019 розничным или корпоративным, используйте формат версии и номера сборки Office. (Для Office 2013 и Office 2021 разница между корпоративной лицензией и розничной лицензией не имеет значения.)

- **Розничная торговля**. Как для Office 2016, так и для 2019 формат имеет формат , оканчивающийся `YYMM (xxxxx.xxxxxx)`двумя блоками из пяти цифр, например `2206 (Build 15330.20264`.
- **С корпоративной лицензией**:
  - Для Office 2016 формат имеет формат , заканчивающийся `16.0.xxxx.xxxxx`двумя блоками из *четырех* цифр, например `16.0.5197.1000`.
  - Для Office 2019 формат имеет формат , заканчивающийся `1808 (xxxxx.xxxxxx)`двумя блоками из *пяти* цифр, например `1808 (Build 10388.20027)`. Обратите внимание, что год и месяц всегда `1808`являются .

| Версия Office | Розничная и корпоративная лицензия | Edge WebView2 (на основе Chromium) установлен? | Браузер |
|:-----|:-----|:-----|:-----|
| Office 2013 | Всё равно | Всё равно | Internet Explorer 11 |
| Office 2016 | Корпоративная лицензия | Всё равно | Internet Explorer 11 |
| Office 2019 | Корпоративная лицензия | Всё равно | Internet Explorer 11 |
| Office 2016 — Office 2019 | Розничная торговля | Нет | Microsoft Edge<sup>1, 2</sup> с оригинальным WebView (EdgeHTML)</br>Если Edge не установлен, используется Internet Explorer 11. |
| Office 2016 — Office 2019 | Розничная торговля | Да<sup>3</sup> | Microsoft Edge<sup>1</sup> с WebView2 (на основе Chromium) |
| Office 2021 | Всё равно | Да<sup>3</sup> | Microsoft Edge<sup>1</sup> с WebView2 (на основе Chromium) |

<sup>1</sup> При использовании Microsoft Edge экранный диктор Windows (иногда называемый "средством чтения с экрана") считывает `<title>` тег на странице, открывающейся в области задач. В Internet Explorer 11 экранный диктор считывает строку заголовка области задач, которая поступает из **\<DisplayName\>** значения манифеста надстройки.

<sup>2</sup> Если надстройка **\<Runtimes\>** содержит элемент в манифесте, она не будет использовать Microsoft Edge с исходным WebView (EdgeHTML). Если выполнены условия использования Microsoft Edge с WebView2 (на основе Chromium), надстройка использует этот браузер. В противном случае используется Internet Explorer 11. Дополнительные сведения см. в статье [Runtimes](/javascript/api/manifest/runtimes).

<sup>3</sup> В версиях Windows до Windows 11 необходимо установить элемент управления WebView2, чтобы Office смог внедрить его. Он устанавливается с бессрочной Office 2021 или более поздней версии, но не устанавливается автоматически с Microsoft Edge. Если у вас есть более ранняя версия Office с бессрочной лицензией, используйте инструкции по установке элемента управления [в Microsoft Edge WebView2 / Внедрение веб-содержимого ... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## <a name="microsoft-365-subscription-versions-of-office-on-windows"></a>Версии Office для Windows по подписке на Microsoft 365

Для Office в Windows по подписке используемый браузер определяется операционной системой, версией Office и установкой Edge WebView2 (на основе Chromium).

|ОС|Версия Office|Edge WebView2 (на основе Chromium) установлен?|Браузер|
|:-----|:-----|:-----|:-----|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp; 16.0.11629<sup>2</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.11629&nbsp;_И_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| Всё равно|Microsoft Edge<sup>1, 3</sup> с оригинальным WebView (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Окно 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Нет |Microsoft Edge<sup>1, 3</sup> с оригинальным WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Да<sup>4</sup>|  Microsoft Edge<sup>1</sup> с WebView2 (на основе Chromium) |

<sup>1</sup> При использовании Microsoft Edge экранный диктор Windows (иногда называемый "средством чтения с экрана") считывает `<title>` тег на странице, открывающейся в области задач. В Internet Explorer 11 экранный диктор считывает строку заголовка области задач, которая поступает из **\<DisplayName\>** значения манифеста надстройки.

<sup>2</sup> Дополнительные сведения см. на [странице журнала обновлений](/officeupdates/update-history-office365-proplus-by-date) и о том, как [найти версию клиента Office и канал обновления](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) .

<sup>3</sup> Если надстройка **\<Runtimes\>** содержит элемент в манифесте, она не будет использовать Microsoft Edge с исходным WebView (EdgeHTML). Если выполнены условия использования Microsoft Edge с WebView2 (на основе Chromium), надстройка использует этот браузер. В противном случае он использует Internet Explorer 11 независимо от версии Windows или Microsoft 365. Дополнительные сведения см. в статье [Runtimes](/javascript/api/manifest/runtimes).

<sup>4</sup> В версиях Windows до Windows 11 необходимо установить элемент управления WebView2, чтобы Office смог внедрить его. Он устанавливается вместе с Microsoft 365 версии 2101 или более поздней, но не устанавливается автоматически вместе с Microsoft Edge. Если у вас более ранняя версия Microsoft 365, используйте инструкции по установке элемента управления [в Microsoft Edge WebView2 / Внедрение веб-содержимого ... с Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). В сборках Microsoft 365 до версии 16.0.14326.xxxxx необходимо также создать раздел реестра **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** и присвоить `dword:00000001`ей значение .

## <a name="working-with-internet-explorer"></a>Работа с Internet Explorer

Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если у любого из пользователей надстройки есть платформы, использующие Internet Explorer 11, то использовать синтаксис и функции ECMAScript 2015 или более поздней версии можно двумя способами.

- Напишите код в ECMAScript 2015 (также называемом ES6) или более поздней версии JavaScript или в TypeScript, а затем скомпилируйте код в ES5 JavaScript с помощью компилятора, например [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
- Напишите в ECMAScript 2015 или более поздней версии JavaScript, но также загрузите библиотеку [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) , например [core-js](https://github.com/zloirock/core-js) , которая позволяет IE выполнять код.

Дополнительные сведения об этих параметрах см. в разделе [Поддержка Internet Explorer 11](../develop/support-ie-11.md).

Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение. Дополнительные сведения см [. в статье Определение надстройки во время выполнения в Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="troubleshoot-microsoft-edge-issues"></a>Устранение неполадок с Microsoft Edge

### <a name="service-workers-are-not-working"></a>Рабочие службы не работают

Надстройки Office не поддерживают рабочие роли службы, если используется исходный microsoft Edge WebView [, EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML). Они поддерживаются в [Chromium Edge WebView2](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) со значением `scrollbar`.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Получение ошибок при попытке загрузить PDF-файл

Прямое скачивание больших двоичных объектов в виде PDF-файлов в надстройке не поддерживается, если браузером является Edge. Обходной путь — создать простое веб-приложение, которое скачивает BLOB-объекты в виде PDF-файлов. В надстройке `Office.context.ui.openBrowserWindow(url)` вызовите метод и передайте URL-адрес веб-приложения. Откроется веб-приложение в окне браузера за пределами Office.

## <a name="wip-protected-documents"></a>Документы, защищенные WIP

Надстройки, работающие в [документе, защищенном WIP,](/windows/uwp/enterprise/wip-hub) никогда не используют **Microsoft Edge с WebView2 (на основе Chromium).** В разделах [Бессрочные версии Office для Windows](#perpetual-versions-of-office-on-windows) и [версии Office 365 с подпиской на Microsoft 365 в Windows, приведенных](#microsoft-365-subscription-versions-of-office-on-windows) выше в этой статье, замените **Microsoft Edge оригинальным WebView (EdgeHTML)** для **Microsoft Edge с WebView2 (на основе Chromium),** где бы ни отображалось последнее.

Чтобы определить, защищен ли документ WIP, выполните следующие действия:

1. Откройте файл.
1. Перейдите на вкладку **Файл** на ленте.
1. Выберите **Сведения**.
1. В левом верхнем углу страницы **Сведений** сразу под именем файла документ с поддержкой WIP будет иметь значок портфеля, за которым следует **управляемый по труду (...).**

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
- [Среды выполнения в надстройках Office](../testing/runtimes.md)
