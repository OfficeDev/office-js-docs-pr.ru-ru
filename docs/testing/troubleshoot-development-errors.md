---
title: Устранение ошибок разработки надстроек Office
description: Узнайте, как устранять ошибки разработки в надстройках Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409418"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Устранение ошибок разработки надстроек Office

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки

Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются

Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере. 

#### <a name="for-windows"></a>Для Windows:

Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` и удалите содержимое папки `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` , если она существует.

#### <a name="for-mac"></a>Для Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Для iOS
Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Изменения статических файлов, таких как JavaScript, HTML и CSS, не отображаются.

Браузер может кэшировать эти файлы. Чтобы избежать этого, отключите кэширование на стороне клиента при разработке. Сведения будут зависеть от того, какой тип сервера вы используете. В большинстве случаев необходимо добавить определенные заголовки в HTTP-ответы. Мы предлагаем следующий набор заголовков:

- Cache-Control: "private, no-cache, no-store"
- Pragma: "No-cache"
- Expires: "-1"

Пример использования на сервере Node.JS Express см. в [этом файле app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js). Пример использования в проекте ASP.NET см. в [этом файле cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

Если надстройка размещена на сервере Internet Information Server (IIS), можно также добавить указанные сведения в файл web.config.

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

Если сначала эти действия безуспешны, вам, возможно, потребуется очистить кэш браузера. Сделайте это в интерфейсе браузера. Иногда очистить кэш браузера Microsoft Edge, используя пользовательский интерфейс, не удается. В таком случае выполните следующую команду в командной строке Windows.

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Изменения, внесенные в значения свойств, не происходят и сообщение об ошибке не отображается

Ознакомьтесь с справочной документацией по свойству, чтобы проверить, доступно ли оно только для чтения. Кроме того, [определения TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта доступны только для чтения. Если вы попытаетесь установить свойство, доступное только для чтения, операция записи завершится с ошибкой без уведомления и не выдается сообщение об ошибке. В следующем примере ошибочно попытаются задать свойство, доступное только для чтения, [Chart.ID](/javascript/api/excel/excel.chart#id). Просмотрите также, что [некоторые свойства не могут быть установлены напрямую](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Надстройка не работает на пограничной стороне, но работает в других браузерах

Ознакомьтесь с [разрешениями проблем Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Надстройка Excel вызывает ошибки, но не всегда

Возможные причины [: Устранение неполадок](../excel/excel-add-ins-troubleshooting.md) в надстройках Excel.

## <a name="see-also"></a>См. также

- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md)
- [Загрузка неопубликованной надстройки Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](testing-and-troubleshooting.md)
