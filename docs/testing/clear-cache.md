---
title: Очистка кэша Office
description: Узнайте, как очищать кэш Office на компьютере.
ms.date: 08/02/2021
ms.localizationpriority: high
ms.openlocfilehash: 4d5351e9f8758109bfd0ef4a901c5ef916c98fa4
ms.sourcegitcommit: e4d98eb90e516b9c90e3832f3212caf48691acf6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2021
ms.locfileid: "60537648"
---
# <a name="clear-the-office-cache"></a>Очистка кэша Office

Чтобы удалить неопубликованную надстройку, ранее установленную в Windows, на компьютерах Mac или в iOS, очистите кэш Office на компьютере.

Кроме того, если вы изменяете манифест надстройки (например, обновляете имена файлов значков или текст команд надстройки), следует очистить кэш Office, а потом выполнить повторную загрузку неопубликованной надстройки с помощью обновленного манифеста. Это позволяет Office отобразить надстройку в соответствии с описанием в обновленном манифесте.

> [!NOTE]
> Для удаления загруженной неопубликованной надстройки из Excel, OneNote, PowerPoint или Word в Интернете см. статью [Загрузка неопубликованных надстроек Office для тестирования в Office для Интернета: удаление загруженной неопубликованной надстройки](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in).

## <a name="clear-the-office-cache-on-windows"></a>Очистка кэша Office в Windows

Чтобы удалить все неопубликованные надстройки из Excel, Word и PowerPoint, удалите содержимое папки.

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Если указанная ниже папка существует, также удалите ее содержимое.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

Чтобы удалить неопубликованную надстройку из Outlook, выполните действия, описанные в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md), чтобы найти надстройку в разделе **Настраиваемые надстройки** диалогового окна, в котором перечислены установленные надстройки. Щелкните многоточие (`...`) для надстройки, а затем выберите **Удалить**, чтобы удалить определенную надстройку. Если такой способ удаления надстроек не работает, удалите содержимое папки `Wef`, как указано выше для Excel, Word и PowerPoint.

Чтобы очистить кэш в Office на Windows 10, когда надстройка работает в Microsoft Edge, вы можете использовать Microsoft Edge DevTools.

> [!TIP]
> Если вы хотите, чтобы в неопубликованной надстройке отражались только последние изменения ее исходных файлов HTML или JavaScript, не нужно очищать кэш. Вместо этого просто переместите фокус в область задач надстройки (щелкнув в любом месте области задач) и нажмите клавиши **CTRL + F5**, чтобы перезагрузить надстройку.

> [!NOTE]
> Для очистки кэша Office с помощью перечисленных ниже действий в вашей надстройке должна быть область задач. Если в вашей надстройке нет пользовательского интерфейса (например, она использует функцию [проверки при отправке](../outlook/outlook-on-send-addins.md)), потребуется добавить в надстройку область задач, использующую такой же домен для [SourceLocation](../reference/manifest/sourcelocation.md), прежде чем можно будет использовать указанные ниже действия для очистки кэша.

1. Установите [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Откройте надстройку в клиенте Office.

3. Запустите Microsoft Edge DevTools.

4. В Microsoft Edge DevTools перейдите на вкладку **Локальные**. Имя вашей надстройки будет указано в списке.

5. Выберите имя надстройки, чтобы присоединить отладчик к надстройке. Откроется новое окно Microsoft Edge DevTools, когда отладчик присоединяется к надстройке.

6. На вкладке **Сеть** в новом окне нажмите **Очистить кэш**.

    ![Снимок экрана Microsoft Edge DevTools с выделенной кнопкой "Очистить кэш"](../images/edge-devtools-clear-cache.png)

7. Если эти действия не привели к нужному результату, попробуйте нажать **Всегда обновлять с сервера**.

    ![Снимок экрана Microsoft Edge DevTools с выделенной кнопкой "Всегда обновлять с сервера"](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Очистка кэша Office на компьютерах Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>Очистка кэша Office в iOS

Чтобы очистить кэш Office в iOS, вызовите `window.location.reload(true)` в JavaScript в надстройке, чтобы запустить принудительную перезагрузку. Также можно переустановить Office.

## <a name="see-also"></a>Дополнительные материалы

- [Устранение ошибок разработки в надстройках Office](troubleshoot-development-errors.md)
- [Отладка надстроек с помощью средств разработчика для Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Отладка надстроек с помощью средств разработчика для устаревшей версии Microsoft Edge](debug-add-ins-using-devtools-edge-legacy.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Загрузка неопубликованных надстроек Office для тестирования](sideload-office-add-ins-for-testing.md)
- [XML-манифест надстроек Office](../develop/add-in-manifests.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
