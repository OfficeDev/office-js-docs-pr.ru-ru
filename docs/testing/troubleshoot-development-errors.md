---
title: Устранение ошибок разработки в надстройках Office
description: Узнайте, как устранять ошибки разработки в надстройки Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 427d35d49339c1130733a3b33aa1bfedc1bd8317
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660076"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Устранение ошибок разработки в надстройках Office

Ниже приведен список распространенных проблем, которые могут возникнуть при разработке надстройки Office.

> [!TIP]
> Очистка кэша Office часто устраняет проблемы, связанные с устаревшим кодом. Это гарантирует отправку последнего манифеста с использованием текущих имен файлов, текста меню и других элементов команды. Дополнительные сведения см [. в статье "Очистка кэша Office"](clear-cache.md).

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки

Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются

Очистка кэша помогает обеспечить использование последней версии манифеста надстройки. Чтобы очистить кэш Office, следуйте инструкциям в разделе ["Очистка кэша Office"](clear-cache.md). Если вы используете Office в Интернете, очистите кэш браузера через пользовательский интерфейс браузера.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Изменения статических файлов, таких как JavaScript, HTML и CSS, не отображаются.

Браузер может кэшировать эти файлы. Чтобы избежать этого, отключите кэширование на стороне клиента при разработке. Сведения будут зависеть от того, какой тип сервера вы используете. В большинстве случаев необходимо добавить определенные заголовки в HTTP-ответы. Мы предлагаем следующий набор.

- Cache-Control: "private, no-cache, no-store"
- Pragma: "No-cache"
- Expires: "-1"

Пример использования на сервере Node.JS Express см. в [этом файле app.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js). Пример использования в проекте ASP.NET см. в [этом файле cshtml](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Изменения, внесенные в значения свойств, не происходят, и сообщение об ошибке отсутствует

Проверьте справочную документацию по свойству, чтобы узнать, доступно ли оно только для чтения. Кроме того, определения [TypeScript для](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) Office JS указывают, какие свойства объекта доступны только для чтения. При попытке задать свойство только для чтения операция записи завершится сбоем без ошибок. В следующем примере ошибочно предпринимается попытка задать свойство только [для чтения Chart.id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member). См. также [, что некоторые свойства нельзя задать напрямую](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Получение ошибки: "Эта надстройка больше недоступна"

Ниже приведены некоторые причины этой ошибки. Если вы обнаруживаете дополнительные причины, сообщите нам об этом с помощью средства обратной связи в нижней части страницы.

- При использовании Visual Studio может возникнуть проблема с загрузкой неопубликованных приложений. Закройте все экземпляры ведущего приложения Office и Visual Studio. Перезапустите Visual Studio и повторите попытку нажатия клавиши F5.
- Манифест надстройки удален из расположения развертывания, например централизованного развертывания, каталога SharePoint или сетевой папки.
- Значение элемента [ID в](/javascript/api/manifest/id) манифесте было изменено непосредственно в развернутой копии. Если по какой-либо причине вы хотите изменить этот идентификатор, сначала удалите надстройку с узла Office, а затем замените исходный манифест измененным манифестом. Многим необходимо очистить кэш Office, чтобы удалить все трассировки исходного файла. Инструкции [по очистке](clear-cache.md) кэша для операционной системы см. в статье "Очистка кэша Office".
- `resid` Манифест надстройки содержит манифест, который не определен ни в разделе "Ресурсы" [](/javascript/api/manifest/resources) манифеста, `resid` **\<Resources\>** либо имеется несоответствие в орфографии между тем, где она используется и где она определена в разделе.
- В манифесте есть `resid` атрибут, который содержит более 32 символов. Атрибут `resid` и атрибут соответствующего `id` **\<Resources\>** ресурса в разделе не могут содержать более 32 символов.
- Надстройка имеет пользовательскую команду надстройки, но вы пытаетесь запустить ее на платформе, которая не поддерживает их. Дополнительные сведения см. в [разделе наборов обязательных элементов команд надстройки](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets).

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Надстройка не работает в Edge, но работает в других браузерах

См [. сведения об устранении неполадок с Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Надстройка Excel выдает ошибки, но не постоянно

Возможные причины см. в статье "Устранение [неполадок надстроек](../excel/excel-add-ins-troubleshooting.md) Excel".

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Ошибки проверки схемы манифеста в проектах Visual Studio

Если вы используете новые функции, для которых требуются изменения файла манифеста, в Visual Studio могут возникнуть ошибки проверки. Например, при добавлении элемента **\<Runtimes\>** для реализации общей среды выполнения JavaScript может появиться следующую ошибку проверки.

**Элемент Host в пространстве имен "http://schemas.microsoft.com/office/taskpaneappversionoverrides" содержит недопустимый дочерний элемент Runtimes в пространстве имен "http://schemas.microsoft.com/office/taskpaneappversionoverrides"**

В этом случае можно обновить XSD-файлы, которые Visual Studio использует, до последних версий. Последние версии схемы находятся [в [MS-URLMXML]: приложение A: полная схема XML](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Поиск XSD-файлов

1. Откройте проект в Visual Studio.
1. В **Обозреватель решений** откройте manifest.xml файла. Манифест обычно находится в первом проекте в вашем решении.
1. Выберите **окно "** > **Свойства представления"** (F4).
1. В **окне свойств нажмите** кнопку с многоточием (...), чтобы открыть редактор **XML-схем** . Здесь можно найти точное расположение папки всех файлов схемы, которые использует проект.

### <a name="update-the-xsd-files"></a>Обновление XSD-файлов

1. Откройте XSD-файл, который требуется обновить, в текстовом редакторе. Имя схемы из ошибки проверки будет коррелировать с именем XSD-файла. Например, откройте **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Найдите обновленную схему [по адресу [MS-URLMXML]: приложение A. Полная схема XML](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Например, TaskPaneAppVersionOverridesV1_0 находится в [схеме taskpaneappversionoverrides](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Скопируйте текст в текстовый редактор.
1. Сохраните обновленный XSD-файл.
1. Перезапустите Visual Studio, чтобы получить новые изменения XSD-файла.

Предыдущий процесс можно повторить для любых дополнительных схем, которые устарели.

## <a name="when-working-offline-no-office-apis-work"></a>При работе в автономном режиме интерфейсы API Office не работают

При загрузке библиотеки JavaScript для Office из локальной копии, а не из CDN API могут перестать работать, если библиотека не является актуальной. Если вы некоторое время не были в проекте, переустановите библиотеку, чтобы получить последнюю версию. Процесс зависит от интегрированной среды разработки. Выберите один из следующих вариантов в зависимости от среды.

- **Visual Studio**: см [. обновление до последней библиотеки API JavaScript для Office](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 
- **Любая другая интегрированная** среда разработки. См. пакеты npm @microsoft [/office-js](https://www.npmjs.com/package/@microsoft/office-js) и [@types/office-js](https://www.npmjs.com/package/@types/office-js).

## <a name="see-also"></a>См. также

- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md)
- [Загрузка неопубликованной надстройки Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Отладка надстроек Office на Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
