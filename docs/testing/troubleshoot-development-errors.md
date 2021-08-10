---
title: Устранение ошибок разработки с Office надстройки
description: Узнайте, как устранить ошибки разработки в Office надстройки.
ms.date: 06/11/2021
localization_priority: Normal
ms.openlocfilehash: a750f8db6e58406403d8bd0ef89e60128c2e08523375b4b2fbe6a904bfbae2d4
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093228"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Устранение ошибок разработки с Office надстройки

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки

Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются

Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере. 

#### <a name="for-windows"></a>Для Windows:

Удалите содержимое папки и удалите содержимое папки, если `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` она существует.

#### <a name="for-mac"></a>Для Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Для iOS

Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Изменения статических файлов, таких как JavaScript, HTML и CSS, не отображаются.

Браузер может кэшировать эти файлы. Чтобы избежать этого, отключите кэширование на стороне клиента при разработке. Сведения будут зависеть от того, какой тип сервера вы используете. В большинстве случаев необходимо добавить определенные заголовки в HTTP-ответы. Предлагаем следующий набор.

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Изменения, внесенные в значения свойств, не происходят, и сообщение об ошибке не сообщается

Проверьте справочную документацию для свойства, чтобы узнать, читается ли оно только. Кроме того, [определения TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объектов являются только для чтения. Если вы попытайтесь установить свойство только для чтения, операция записи не будет работать без ошибки. В следующем примере ошибочно пытается установить свойство только [для чтения Chart.id](/javascript/api/excel/excel.chart#id). См. [также Некоторые свойства не могут быть установлены напрямую](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Получение ошибки: "Эта надстройка больше недоступна"

Ниже приводится несколько причин этой ошибки. Если вы обнаружите дополнительные причины, сообщите нам с помощью средства обратной связи в нижней части страницы.

- Если вы используете Visual Studio, может возникнуть проблема с боковой загрузкой. Закрой все экземпляры Office и Visual Studio. Перезапустите Visual Studio и повторите нажатие F5.
- Манифест надстройки удален из расположения развертывания, например централизированного развертывания, каталога SharePoint или сетевой доли.
- Значение элемента [ID](../reference/manifest/id.md) в манифесте было изменено непосредственно в развернутой копии. Если по какой-либо причине необходимо изменить этот ID, сначала удалите надстройки из Office, а затем замените исходный манифест на измененный манифест. Многим требуется очистить кэш Office, чтобы удалить все следы оригинала. См. раздел Изменения в командах [надстройки,](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) включая кнопки ленты и элементы меню, не вступает в силу ранее в этой статье.
- Манифест надстройки имеет манифест, который не определен нигде в разделе Ресурсы манифеста, или существует несоответствие в написании между тем, где он используется и где он определен в `resid` [](../reference/manifest/resources.md) `resid` `<Resources>` разделе.
- В манифесте есть атрибут с более `resid` чем 32 символами. Атрибут и атрибут соответствующего ресурса в разделе не могут быть более `resid` `id` `<Resources>` 32 символов.
- Надстройка имеет настраиваемую команду надстройки, но вы пытаетесь запустить ее на платформе, которая не поддерживает их. Дополнительные сведения см. в [дополнительных наборах требований к командам надстройки.](../reference/requirement-sets/add-in-commands-requirement-sets.md)

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Надстройка не работает в Edge, но работает в других браузерах

См. [в Microsoft Edge устранение неполадок.](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel надстройка бросает ошибки, но не последовательно

См. [Excel возможные](../excel/excel-add-ins-troubleshooting.md) причины устранения неполадок.

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Ошибки проверки схемы манифеста в Visual Studio проектах

Если вы используете новые функции, которые требуют изменений в файл манифеста, вы можете получить ошибки проверки в Visual Studio. Например, при добавлении элемента для реализации общего времени выполнения JavaScript вы можете увидеть `<Runtimes>` следующую ошибку проверки.

**Элемент "Host" в пространстве имен ' имеет недействительный детский элемент http://schemas.microsoft.com/office/taskpaneappversionoverrides 'Runtimes' в пространстве имен http://schemas.microsoft.com/office/taskpaneappversionoverrides '**

В этом случае можно обновить XSD-файлы, Visual Studio используются в последних версиях. Последние версии схемы находятся в [[MS-OWEMXML]: Приложение A: Полная схема XML](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Найдите XSD-файлы

1. Откройте проект в Visual Studio.
1. В **обозревателе** решений откройте manifest.xml файл. Манифест обычно находится в первом проекте под вашим решением.
1. Выберите **окно Свойства**  >  **представления** (F4).
1. В **окне Свойства** выберите ellipsis (...) для открытия **редактора схем XML.** Здесь вы можете найти точное расположение папок всех файлов схемы, которые использует проект.

### <a name="update-the-xsd-files"></a>Обновление XSD-файлов

1. Откройте XSD-файл, который необходимо обновить в текстовом редакторе. Имя схемы из ошибки проверки будет соотноситься с именем файла XSD. Например, откройте **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Найдите обновленную схему [в [MS-OWEMXML]: Приложение A: Полная схема XML](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Например, TaskPaneAppVersionOverridesV1_0 [в taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Скопируйте текст в текстовый редактор.
1. Сохраните обновленный XSD-файл.
1. Перезапустите Visual Studio, чтобы получить новые изменения XSD-файла.

Вы можете повторить предыдущий процесс для любых дополнительных схем, которые устарели.

## <a name="see-also"></a>См. также

- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md)
- [Загрузка неопубликованной надстройки Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](testing-and-troubleshooting.md)
