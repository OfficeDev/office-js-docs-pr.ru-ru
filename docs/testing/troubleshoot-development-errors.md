---
title: Устранение ошибок разработки с помощью надстройки Office
description: Узнайте, как устранять ошибки разработки в надстройки Office.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 48216230db4bf90ca53ef10d98786877bd3905c2
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771426"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Устранение ошибок разработки с помощью надстройки Office

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки

Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются

Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере. 

#### <a name="for-windows"></a>Для Windows:

Удалите содержимое папки и удалите ее( `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` если `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` она существует).

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Изменения значений свойств не происходят, и сообщение об ошибке не сообщается

Проверьте справочную документацию по свойству, чтобы узнать, прочитано ли оно только. Кроме того, определения [TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта являются только для чтения. Если попытаться установить свойство только для чтения, операция записи будет неудачной без ошибок. В следующем примере ошибочно предпринимается попытка установить свойство только [для Chart.id.](/javascript/api/excel/excel.chart#id) См. [также, что некоторые свойства нельзя настроить напрямую.](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Ошибка при получении: "Эта надстройка больше недоступна"

Ниже следующую часть причин этой ошибки. Если вы обнаружите дополнительные причины, сообщите нам с помощью средства обратной связи в нижней части страницы.

- Если вы используете Visual Studio, возможно, возникла проблема с загрузкой неогрузки. Закроем все экземпляры ведущего экземпляра Office и Visual Studio. Перезапустите Visual Studio и повторите нажатие F5.
- Манифест надстройки удален из расположения развертывания, например централизованного развертывания, каталога SharePoint или сетевой сети.
- Значение элемента [ID](../reference/manifest/id.md) в манифесте было изменено непосредственно в развернутой копии. Если по какой-либо причине вы хотите изменить этот ИД, сначала удалите надстройки из ведущего office, а затем замените исходный манифест на измененный манифест. Многим необходимо очистить кэш Office, чтобы удалить все трассировки исходного. См. раздел "Изменения команд [надстройки",](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) включая кнопки ленты и пункты меню, которые не вступили в силу ранее в этой статье.
- Манифест надстройки имеет манифест, который не определен ни в разделе "Ресурсы" манифеста, либо имеется несоответствие в орфографии между местом ее использования и местом, где оно определено в `resid` [](../reference/manifest/resources.md) `resid` `<Resources>` разделе.
- В `resid` манифесте есть атрибут, в который вмеется более 32 символов. Атрибут и атрибут соответствующего ресурса в разделе не могут быть больше `resid` `id` `<Resources>` 32 символов.
- Надстройка имеет пользовательскую команду надстройки, но вы пытаетесь запустить ее на платформе, которая их не поддерживает. Дополнительные сведения см. в наборах требований [для команд надстройки.](../reference/requirement-sets/add-in-commands-requirement-sets.md)

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Надстройка не работает в Edge, но работает в других браузерах

См. [устранение неполадок Microsoft Edge.](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Надстройка Excel высылает ошибки, но не постоянно

Возможные причины см. в устранении [неполадок](../excel/excel-add-ins-troubleshooting.md) надстройки Excel.

## <a name="see-also"></a>См. также

- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md)
- [Загрузка неопубликованной надстройки Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](testing-and-troubleshooting.md)
