---
title: Браузеры, используемые надстройками Office
description: Указывается, как операционная система и версия Office определяют браузер, используемый надстройками Office.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 37d001d7feb170b11edc4f6a233f6fdc15cf3438
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950364"
---
# <a name="browsers-used-by-office-add-ins"></a>Браузеры, используемые надстройками Office

Надстройки Office — это веб-приложения, отображаемые с помощью элементов iFrame при работе в Office в Интернете и с помощью встроенных элементов управления браузеров в Office для настольных и мобильных клиентов. Для запуска JavaScript надстройкам также требуется модуль JavaScript. Встроенный браузер и модуль поставляются браузером, установленным на компьютере пользователя.

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
|Windows 10 версии 1903 или выше / Office 365 версии ниже 16.0.11629|Internet Explorer 11|
|Windows 10 версии 1903 или выше / Office 365 версии 16.0.11629 или выше|Microsoft Edge\*|

\* Если используется Microsoft Edge, экранный диктор Windows 10 (его иногда называют "читатель экрана") считывает тег `<title>` на странице, которая открывается в области задач. Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если какой-либо пользователь вашей надстройки применяет платформы с Internet Explorer 11, для применения синтаксиса и возможностей ECMAScript 2015 или более поздних версий вам нужно либо транскомпилировать свой код JavaScript в ES5, либо использовать полизаполнение. Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

## <a name="troubleshooting-microsoft-edge-issues"></a>Устранение проблем с Microsoft Edge

### <a name="chromium-based-edge-is-installed-on-my-development-computer-but-my-add-in-does-not-use-it"></a>На моем компьютере разработчика установлен граничный сервер чромиум, но надстройка не использует ее

Базовый браузер в [Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium) изменился на чромиум. Старая база, называемая Еджехтмл, не удаляется при установке пограничного сервера на основе Чромиум. Office по-прежнему будет использовать базу Еджехтмл для надстроек до тех пор, пока не будет установлена сборка Office 365, поддерживающая Чромиум на компьютере. Мы ожидаем, что эти сборки поставляются в 2020. Скорее всего, они будут отображаться в канале "предварительные сотрудники" в первой половине года.

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>В области задач не отображается полоса прокрутки

По умолчанию полосы прокрутки в Microsoft Edge скрыты до наведения указателя мыши. Чтобы полоса прокрутки отображалась постоянно, стиль CSS, применяемый к элементу `<body>` страниц в области задач, должен содержать свойство [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) со значением `scrollbar`. 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>При отладке с помощью Microsoft Edge DevTools надстройка аварийно завершает работу или перезагружается

Настроенные точки останова в [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) приложение Office может воспринимать как зависание надстройки. В этом случае выполняется автоматическая перезагрузка надстройки. Чтобы избежать этого, добавьте следующий раздел реестра и значение на компьютере разработчика: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>При попытке открытия надстройки появляется сообщение "ОШИБКА НАДСТРОЙКИ. Не удается открыть эту надстройку из localhost"

Одной из известных причин является требование Microsoft Edge, чтобы для localhost предоставлялось исключение замыкания на себя. Следуйте инструкциям из статьи [Не удается открыть надстройку из localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).


## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](requirements-for-running-office-add-ins.md)
