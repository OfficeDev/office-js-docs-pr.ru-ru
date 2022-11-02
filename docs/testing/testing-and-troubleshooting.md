---
title: Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office
description: Узнайте, как устранять ошибки пользователей в надстройках Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18bb3c180cd3af1eb8d045d7c69b9772532b04d4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810374"
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
|Ошибка: объект не поддерживает свойство или метод 'defineProperty'|Убедитесь, что Internet Explorer не работает в режиме совместимости. Перейдите в раздел **Средства** > **Параметры представления совместимости**.|
|К сожалению, не удалось загрузить приложение, так как ваша версия браузера не поддерживается. Чтобы открыть список поддерживаемых версий браузеров, щелкните здесь.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>При установке надстройки в строке состояния появляется сообщение "Ошибка при загрузке надстройки"

1. Закройте Office.
1. Убедитесь, что манифест является допустимым. См [. раздел Проверка манифеста надстройки Office](troubleshoot-manifest.md).
1. Перезапустить надстройку.
1. Переустановите надстройку.

Также можно отправить нам отзыв: при использовании Excel для Windows или Mac можно отправить отзыв группе расширяемости Office непосредственно из Excel. Для этого выберите **Файл** > **Отзывы и предложения** > **Отправить нахмуренный смайлик**. При отправке нахмуренного смайлика будут предоставлены необходимые журналы для понимания описываемой проблемы.

## <a name="outlook-add-in-doesnt-work-correctly"></a>Надстройка Outlook работает неправильно

Если надстройка Outlook в Windows и [в Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) работает неправильно, попробуйте включить отладку сценариев в Internet Explorer.

- Перейдите в **раздел Инструменты** > ,  > **свойства браузера****Дополнительно**.
- В разделе **Обзор**, снимите флажки **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.

We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.

## <a name="add-in-doesnt-activate-in-office-2013"></a>Надстройка не активируется в Office 2013

Если надстройка не активируется, когда пользователь выполняет следующие действия.

1. выполнении входа с помощью учетной записи Майкрософт в Office 2013;

1. включении двухшаговой проверки учетной записи Майкрософт;

1. проверки своего удостоверения по запросу при попытке добавления надстройки, —

убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).

## <a name="add-in-dialog-box-cannot-be-displayed"></a>Не отображается диалоговое окно надстройки

При открытии надстройки Office пользователю будет предложено разрешить отображение диалогового окна. Пользователь выбирает **Разрешить**, и появляется следующее сообщение об ошибке.

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![Снимок экрана: сообщение об ошибке диалогового окна.](../images/dialog-prevented.png)

|Браузеры|Платформы|
|:--------------------|:---------------------|
|Microsoft Edge|Office в Интернете|

Чтобы устранить эту проблему, пользователи или администраторы могут добавить домен надстройки в список доверенных сайтов в браузере Microsoft Edge.

> [!IMPORTANT]
> Не добавляйте URL-адрес надстройки в список надежных сайтов, если вы не доверяете надстройке.

Чтобы добавить URL-адрес в список надежных сайтов:

1. На **панели управления** перейдите в раздел **Свойства браузера** > **Безопасность**.
1. Выберите зону **Надежные сайты** и нажмите кнопку **Сайты**.
1. Введите URL-адрес из сообщения об ошибке и нажмите кнопку **Добавить**.
1. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a>Дополнительные материалы

- [Устранение ошибок разработки в надстройках Office](troubleshoot-development-errors.md)
