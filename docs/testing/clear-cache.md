---
title: Очистка кэша Office
description: Узнайте, как очищать кэш Office на компьютере.
ms.date: 11/15/2021
ms.localizationpriority: high
ms.openlocfilehash: 36f3de58eb5089f6c638510cb33879cb36c7330c
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081409"
---
# <a name="clear-the-office-cache"></a>Очистка кэша Office

Чтобы удалить неопубликованную надстройку, ранее установленную в Windows, на компьютерах Mac или в iOS, очистите кэш Office на компьютере.

Кроме того, если вы изменяете манифест надстройки (например, обновляете имена файлов значков или текст команд надстройки), следует очистить кэш Office, а потом выполнить повторную загрузку неопубликованной надстройки с помощью обновленного манифеста. Это позволяет Office отобразить надстройку в соответствии с описанием в обновленном манифесте.

> [!NOTE]
> Для удаления загруженной неопубликованной надстройки из Excel, OneNote, PowerPoint или Word в Интернете см. статью [Загрузка неопубликованных надстроек Office для тестирования в Office для Интернета: удаление загруженной неопубликованной надстройки](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in).

## <a name="clear-the-office-cache-on-windows"></a>Очистка кэша Office в Windows

Существует три метода очистки кэша Office на компьютере с Windows: автоматически, вручную и с помощью средств разработчика Microsoft Edge. Эти методы описаны в следующих подразделах.

### <a name="automatically"></a>Автоматически

Этот метод рекомендуется использовать для компьютеров разработки надстройки. Если Office используется в Windows версии 2108 или более поздней, следующие действия настраивают автоматическую очистку кэша Office при каждом повторном открытии Office.

> [!NOTE]
> Автоматический метод не поддерживается для Outlook.

1. На ленте любого ведущего приложения Office (кроме Outlook) выберите **Файл** > **Параметры** > **Центр управления безопасностью** > **Параметры центра управления безопасностью** > **Доверенные каталоги надстроек**.
1. Установите флажок **При следующем запуске Office очистите кэш всех ранее запущенных веб-надстроек**.

### <a name="manually"></a>Вручную

Ручной метод для Excel, Word и PowerPoint отличается от Outlook.

#### <a name="manually-clear-the-cache-in-excel-word-and-powerpoint"></a>Очистка кэша вручную в Excel, Word и PowerPoint

Чтобы удалить все неопубликованные надстройки из Excel, Word и PowerPoint, удалите содержимое следующей папки.

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Если указанная ниже папка существует, также удалите ее содержимое.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

#### <a name="manually-clear-the-cache-in-outlook"></a>Очистка кэша вручную в Outlook

Чтобы удалить неопубликованную надстройку из Outlook, выполните действия, описанные в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md), чтобы найти надстройку в разделе **Настраиваемые надстройки** диалогового окна, в котором перечислены установленные надстройки. Щелкните многоточие (`...`) для надстройки, а затем выберите **Удалить**, чтобы удалить определенную надстройку. Если такой способ удаления надстроек не работает, удалите содержимое папки `Wef`, как указано выше для Excel, Word и PowerPoint.

### <a name="using-the-microsoft-edge-developer-tools"></a>С помощью средств разработчика Microsoft Edge

Чтобы очистить кэш Office в Windows 10, когда надстройка работает в Microsoft Edge, можно использовать средства разработчика Microsoft Edge.

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
- [Отладка надстроек с помощью средств разработчика в Microsoft Edge (на основе Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Загрузка неопубликованных надстроек Office для тестирования](sideload-office-add-ins-for-testing.md)
- [XML-манифест надстроек Office](../develop/add-in-manifests.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
