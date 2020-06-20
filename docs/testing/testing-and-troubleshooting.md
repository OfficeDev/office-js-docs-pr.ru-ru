---
title: Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office
description: Узнайте, как устранять ошибки пользователей в надстройках Office.
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 1dbc8cc18e0c9b12ccff605b655dd7c8629fb9cf
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810851"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in. 

Для выявления и устранения проблем с надстройками также можно использовать [Fiddler](https://www.telerik.com/fiddler).

## <a name="common-errors-and-troubleshooting-steps"></a>Распространенные ошибки и инструкции по устранению неполадок

В таблице ниже перечислены распространенные сообщения об ошибках, с которыми могут столкнуться пользователи, и действия, которые можно предпринять для устранения ошибки.



|**Сообщение об ошибке**|**Решение**|
|:-----|:-----|
|Ошибка приложения: не удалось подключиться к каталогу|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|Убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).|
|Ошибка: объект не поддерживает свойство или метод 'defineProperty'|Убедитесь, что Internet Explorer не работает в режиме совместимости. Откройте меню "Сервис" > **Параметры просмотра в режиме совместимости**.|
|Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>При установке надстройки в строке состояния появляется сообщение "Ошибка при загрузке надстройки"

1. Закройте Office.
2. Убедитесь, что манифест действителен.
3. Перезапустите надстройку.
4. Переустановите надстройку.

Также можно отправить нам отзыв: при использовании Excel для Windows или Mac можно отправить отзыв группе расширяемости Office непосредственно из Excel. Для этого выберите **Файл** | **Отзывы и предложения** | **Отправить нахмуренный смайлик**. При отправке нахмуренного смайлика будут предоставлены необходимые журналы для понимания описываемой проблемы.

## <a name="outlook-add-in-doesnt-work-correctly"></a>Надстройка Outlook работает неправильно

Если надстройка Outlook в Windows и [в Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) работает неправильно, попробуйте включить отладку сценариев в Internet Explorer. 


- Откройте раздел Сервис > " **Свойства браузера**"  >  **Advanced**.
    
- В разделе **Обзор**, снимите флажки **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.
    
We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.


## <a name="add-in-doesnt-activate-in-office-2013"></a>Надстройка не активируется в Office 2013

Если надстройка не активируется при выполнении пользователем следующих действий:


1. выполнении входа с помощью учетной записи Майкрософт в Office 2013;
    
2. включении двухшаговой проверки учетной записи Майкрософт;
    
3. проверки своего удостоверения по запросу при попытке добавления надстройки, —
    
убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки

Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).


## <a name="add-in-dialog-box-cannot-be-displayed"></a>Не отображается диалоговое окно надстройки

When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![Снимок экрана: сообщение об ошибке](http://i.imgur.com/3mqmlgE.png)

|**Браузеры**|**Платформы**|
|:--------------------|:---------------------|
|Internet Explorer, Microsoft Edge|Office в Интернете|

To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.

> [!IMPORTANT]
> Не добавляйте URL-адрес надстройки в список надежных сайтов, если вы не доверяете надстройке.

Чтобы добавить URL-адрес в список надежных сайтов:

1. На **панели управления** перейдите в раздел **Свойства браузера** > **Безопасность**.
2. Выберите зону **Надежные сайты** и нажмите кнопку **Сайты**.
3. Введите URL-адрес из сообщения об ошибке и нажмите кнопку **Добавить**.
4. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются

Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере. 

#### <a name="for-windows"></a>Для Windows:
Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>Для Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Для iOS
Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

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

## <a name="see-also"></a>См. также

- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md) 
- [Загрузка неопубликованной надстройки Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Расширение отладчика надстроек Microsoft Office для Visual Studio Code](./debug-with-vs-extension.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
