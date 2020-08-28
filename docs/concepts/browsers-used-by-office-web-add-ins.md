---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: 4c4462434480a82e69ebaf5a2599ceda1e4693a7
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293060"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, которые отображаются с помощью iFrames при работе в Office в Интернете и с использованием встроенных элементов управления браузером в Office для настольных и мобильных клиентов. Для запуска JavaScript надстройкам также требуется модуль JavaScript. Как встроенный браузер, так и модуль предоставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- , Работает ли надстройка в Office в Интернете, Microsoft 365 или не в подписке Office 2013 или более поздней версии.

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|СОВМЕСТИМ|Версия Office|Установлен пограничный WebView2 (на основе Чромиум)?|Браузер|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|любой|Office в Интернете|Неприменимо|Браузер, в котором открыт Office.|
|Mac|любой|Неприменимо|Safari|
|iOS|любой|Неприменимо|Safari|
|Android|любой|Неприменимо|Chrome|
|Windows 7, 8,1, 10 | не подписка Office 2013 или более поздняя версия|Всё равно|Internet Explorer 11|
|Windows 7 | Microsoft 365| Всё равно | Internet Explorer 11|
|Windows 8,1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Нет| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Всё равно|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _и_ &nbsp; < &nbsp; 16.0.13127.20082<sup>1</sup>| Всё равно|Microsoft Edge<sup>2, 3</sup> с исходным Вебвиев (еджехтмл)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13127.20082<sup>1</sup>| Нет |Microsoft Edge<sup>2, 3</sup> с исходным Вебвиев (еджехтмл)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13127.20082<sup>1</sup>| Да|  Просмотрите Примечание 4. |

<sup>1</sup> ознакомьтесь со [страницей "журнал обновлений](/officeupdates/update-history-office365-proplus-by-date) " и Узнайте, как [найти версию клиента Office и канал обновления](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) для получения дополнительных сведений.

<sup>2</sup> при использовании Microsoft Edge экранный диктор Windows 10 (иногда называется "средство чтения с экрана") считывает `<title>` тег на странице, которая открывается в области задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

<sup>3</sup> если надстройка содержит `Runtimes` элемент в манифесте, он использует Internet Explorer 11 независимо от версии Windows или Microsoft 365. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> браузер, используемый для этой комбинации версий, зависит от канала обновления подписки Microsoft 365. Если пользователь находится на [канале бета-версии](https://insider.office.com/join/windows) (ранее он быстро является быстрым каналом), Office использует Microsoft Edge с WebView2 (чромиум на основе). Для любого другого канала Office использует Microsoft Edge с исходной Вебвиев (Еджехтмл). Поддержка WebView2 в других каналах ожидается на ранних 2021.
> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если какой-либо пользователь вашей надстройки применяет платформы с Internet Explorer 11, для применения синтаксиса и возможностей ECMAScript 2015 или более поздних версий вам нужно либо транскомпилировать свой код JavaScript в ES5, либо использовать полизаполнение. Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение проблем с Microsoft Edge

### <a name="service-workers-are-not-working"></a>Рабочие процессы не работают

Надстройки Office не поддерживают сотрудников службы при использовании исходной [Вебвиев Microsoft Edge](/microsoft-edge/hosting/webview) . Они поддерживаются [пограничным WebView2 на основе чромиум](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) со значением `scrollbar`. 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Получение сообщений об ошибках при попытке загрузить PDF-файл

Непосредственная загрузка больших двоичных объектов как PDF-файлов в надстройке не поддерживается, если пограничный сервер — браузер. Чтобы устранить эту проблемы, создайте простое веб-приложение, которое загружает большие двоичные объекты как PDF-файлы. В надстройке вызовите `Office.context.ui.openBrowserWindow(url)` метод и передайте URL-адрес веб-приложения. Это приведет к открытию веб-приложения в окне браузера вне Office.

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
