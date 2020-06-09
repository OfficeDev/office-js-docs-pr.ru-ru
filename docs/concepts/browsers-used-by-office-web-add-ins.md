---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 2dc66439ff4ab7f9bee148168df4d9d9b30a11a1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608043"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, отображаемые с помощью элементов iFrame при работе в Office в Интернете и с помощью встроенных элементов управления браузеров в Office для настольных и мобильных клиентов. Для запуска JavaScript надстройкам также требуется модуль JavaScript. Как встроенный браузер, так и модуль предоставляются браузером, установленным на компьютере пользователя.

Используемый браузер зависит от указанных ниже факторов.

- Операционная система компьютера.
- Работает надстройка в Office в Интернете, Office 365 или же Office 2013 либо более поздней версии без подписки.

В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.

|**ОС / платформа**|**Browser**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office в Интернете|Браузер, в котором открыт Office.|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows / Office 2013 или более поздней версии без подписки|Internet Explorer 11|
|Windows 10 версии ниже 1903 / Office 365|Internet Explorer 11|
|Windows 10 версии >= 1903/Office 365 ver < 16.0.11629<sup>1</sup>|Internet Explorer 11|
|Windows 10 версии >= 1903/Office 365 ver >= 16.0.11629<sup>1</sup>|Microsoft Edge<sup>2, 3</sup>|

<sup>1</sup> ознакомьтесь со [страницей "журнал обновлений](/officeupdates/update-history-office365-proplus-by-date) " и Узнайте, как [найти версию клиента Office и канал обновления](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) для получения дополнительных сведений.

<sup>2</sup> при использовании Microsoft Edge экранный диктор Windows 10 (иногда называется "средство чтения с экрана") считывает `<title>` тег на странице, которая открывается в области задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

<sup>3</sup> если надстройка содержит `Runtimes` элемент в манифесте, он использует Internet Explorer 11 независимо от версии Windows или Office 365. Дополнительные [сведения см.](../reference/manifest/runtimes.md)

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если какой-либо пользователь вашей надстройки применяет платформы с Internet Explorer 11, для применения синтаксиса и возможностей ECMAScript 2015 или более поздних версий вам нужно либо транскомпилировать свой код JavaScript в ES5, либо использовать полизаполнение. Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение проблем с Microsoft Edge

### <a name="service-workers-are-not-working"></a>Рабочие процессы не работают

Надстройки Office не поддерживают служебных рабочих процессов в [Microsoft Edge вебвиев](/microsoft-edge/hosting/webview). Ознакомьтесь со статьей " [Обзор надстроек Office](../overview/office-add-ins.md) " для получения последних поддерживаемых функций для элемента управления вебвиев Edge. Мы работаем над тем, чтобы создать новую [WebView2 пограничный сервер на основе чромиум](/microsoft-edge/hosting/webview2) на платформе надстроек Office, которые мы планируем поддерживать для сотрудников службы.

### <a name="chromium-based-edge-is-installed-on-my-development-computer-but-my-add-in-does-not-use-it"></a>На моем компьютере разработчика установлен граничный сервер чромиум, но надстройка не использует ее

Базовый браузер в [Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium) изменился на чромиум. Старая база, называемая Еджехтмл, не удаляется при установке пограничного сервера на основе Чромиум. Office по-прежнему будет использовать базу Еджехтмл для надстроек до тех пор, пока не будет установлена сборка Office 365, поддерживающая Чромиум на компьютере. Мы ожидаем, что эти сборки поставляются в 2020. Скорее всего, они будут отображаться в канале "предварительные сотрудники" в первой половине года.

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
