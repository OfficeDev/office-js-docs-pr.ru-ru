---
title: Очистка кэша Office
description: Узнайте, как очищать кэш Office на компьютере.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915080"
---
# <a name="clear-the-office-cache"></a>Очистка кэша Office

Можно удалить надстройку, ранее установленную в Windows, на компьютерах Mac или в iOS, очистив кэш Office на компьютере. 

Кроме того, если вы изменяете манифест надстройки (например, обновляете имена файлов значков или текст команд надстройки), следует очистить кэш Office, а потом заново установить надстройку с помощью обновленного манифеста. В этом случае надстройка будет отображаться в Office в соответствии с обновленным манифестом.

## <a name="clear-the-office-cache-on-windows"></a>Очистка кэша Office в Windows

Чтобы очистить кэш Office в Windows, удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

## <a name="clear-the-office-cache-on-mac"></a>Очистка кэша Office на компьютерах Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a>Очистка кэша Office в iOS

Чтобы очистить кэш Office в iOS, вызовите `window.location.reload(true)` в JavaScript в надстройке, чтобы запустить принудительную перезагрузку. Также можно переустановить Office.

## <a name="see-also"></a>См. также

- [XML-манифест надстройки Office](../develop/add-in-manifests.md)
- [Проверка манифеста надстройки Office](troubleshoot-manifest.md)
- [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md)
- [Загрузка неопубликованных надстроек Office для тестирования](sideload-office-add-ins-for-testing.md)
- [Отладка надстроек Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)